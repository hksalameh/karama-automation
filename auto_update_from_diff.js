const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const xlsx = require('xlsx');

const LOGIN_URL = 'http://92.253.101.105:8040/ICCS/index.aspx';
const DISBURSEMENT_URL_HINT = '/ICCS/Family_sub2.aspx';

// هذه مهلة قصوى فقط. البرنامج لا ينتظرها إذا تحقق الشرط قبل ذلك.
const SITE_ACTION_TIMEOUT_MS = 60000;
const POLL_MS = 100;
const AFTER_ACTION_SETTLE_MS = 120;
const FIELD_SETTLE_MS = 60;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function normalizeText(value) {
  return String(value ?? '')
    .trim()
    .replace(/[أإآ]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ة/g, 'ه')
    .replace(/ـ/g, '')
    .replace(/\s+/g, ' ');
}

function normalizeDigits(value) {
  return String(value ?? '')
    .trim()
    .replace(/[٠-٩]/g, (d) => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))
    .replace(/[^\d]/g, '');
}

function normalizeAmount(value) {
  const s = String(value ?? '')
    .trim()
    .replace(/[٠-٩]/g, (d) => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))
    .replace(/,/g, '')
    .replace(/[^\d.-]/g, '');
  if (!s) return '';
  const n = Number(s);
  return Number.isFinite(n) ? String(n) : s;
}

function outputPath(fileName) {
  const dir = process.env.KARAMA_OUTPUT_DIR || process.cwd();
  fs.mkdirSync(dir, { recursive: true });
  return path.join(dir, fileName);
}

function log(message) {
  console.log(message);
}

function progress(index, total, record, status) {
  const safe = (v) => String(v ?? '').replace(/[|\r\n]/g, ' ');
  console.log(`__PROGRESS__|${index}|${total}|${safe(record.natId)}|${safe(record.oldAmount)}|${safe(record.newAmount)}|${safe(status)}`);
}

async function waitUntil(check, description, timeoutMs = SITE_ACTION_TIMEOUT_MS) {
  const deadline = Date.now() + timeoutMs;
  let lastError = '';
  while (Date.now() < deadline) {
    try {
      if (await check()) return true;
    } catch (err) {
      lastError = err?.message || String(err);
    }
    await sleep(POLL_MS);
  }
  throw new Error(`انتهت مهلة انتظار ${description}${lastError ? `: ${lastError}` : ''}.`);
}

function readConfig() {
  const excelPath = process.env.KARAMA_EXCEL_PATH || '';
  const username = process.env.KARAMA_USERNAME || '';
  const password = process.env.KARAMA_PASSWORD || '';
  const year = String(process.env.KARAMA_YEAR || '').trim();
  const month = String(process.env.KARAMA_MONTH || '').trim();
  const category = String(process.env.KARAMA_CATEGORY || '').trim();

  if (!excelPath || !fs.existsSync(excelPath)) {
    throw new Error(`ملف نتيجة المقارنة غير موجود: ${excelPath || '(فارغ)'}`);
  }
  if (!username || !password) throw new Error('اسم المستخدم أو كلمة المرور غير مدخلين.');
  if (!/^\d{4}$/.test(year)) throw new Error(`السنة غير صحيحة: ${year}`);

  const monthNum = Number(month);
  if (!Number.isInteger(monthNum) || monthNum < 1 || monthNum > 12) {
    throw new Error(`الشهر غير صحيح: ${month}`);
  }
  if (!category) throw new Error('التصنيف غير محدد.');

  return { excelPath, username, password, year, month: String(monthNum), category };
}

function detectHeader(headers, predicates, fallback = null) {
  for (const header of headers) {
    const normalized = normalizeText(header);
    if (predicates.some((p) => p(normalized))) return header;
  }
  return fallback;
}

function loadRecords(excelPath) {
  const wb = xlsx.readFile(excelPath, { cellDates: false });
  let sheetName = wb.SheetNames.find((name) => normalizeText(name).includes('يحتاج تعديل'));
  if (!sheetName) sheetName = wb.SheetNames.find((name) => normalizeText(name).includes('تعديل'));
  if (!sheetName) {
    throw new Error(`لم يتم العثور على ورقة "يحتاج تعديل". الأوراق: ${wb.SheetNames.join(', ')}`);
  }

  const rows = xlsx.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '', raw: false });
  if (!rows.length) return [];

  const headers = Object.keys(rows[0]);
  const natHeader = detectHeader(headers, [(h) => h.includes('وطني')]);
  const nameHeader = detectHeader(headers, [(h) => h.includes('اسم') && h.includes('الموقع')]);
  const oldHeader = detectHeader(headers, [(h) => h.includes('مبلغ') && h.includes('الموقع')]);
  const newHeader = detectHeader(headers, [(h) => h.includes('مبلغ') && (h.includes('فعلي') || h.includes('الجديد') || h.includes('متوقع'))]);
  const reasonHeader = detectHeader(headers, [(h) => h.includes('سبب')]);

  if (!natHeader || !oldHeader || !newHeader) {
    throw new Error(`تعذر اكتشاف أعمدة ملف النتيجة. الأعمدة الموجودة: ${headers.join(' | ')}`);
  }

  return rows
    .map((row) => ({
      natId: normalizeDigits(row[natHeader]),
      name: String(nameHeader ? row[nameHeader] : '').trim(),
      oldAmount: normalizeAmount(row[oldHeader]),
      newAmount: normalizeAmount(row[newHeader]),
      reason: String(reasonHeader ? row[reasonHeader] : '').trim(),
    }))
    .filter((r) => r.natId && r.newAmount !== '');
}

async function launchBrowser() {
  const launchOptions = { headless: false };
  const attempts = [
    ['Chrome', { channel: 'chrome', ...launchOptions }],
    ['Edge', { channel: 'msedge', ...launchOptions }],
    ['Chromium', launchOptions],
  ];

  const errors = [];
  for (const [name, options] of attempts) {
    try {
      const browser = await chromium.launch(options);
      log(`✅ تم تشغيل المتصفح: ${name}`);
      return browser;
    } catch (err) {
      errors.push(`${name}: ${err.message || err}`);
    }
  }
  throw new Error(`تعذر تشغيل Chrome أو Edge.\n${errors.join('\n')}`);
}

async function controlLabel(control) {
  return control.evaluate((el) => `${el.value || ''} ${el.innerText || el.textContent || ''}`.trim()).catch(() => '');
}

function isFinalSaveLabel(label) {
  const normalized = normalizeText(label);
  return normalized.includes('حفظ') && normalized.includes('نهائي');
}

async function waitUntilEnabled(locator, description, timeoutMs = SITE_ACTION_TIMEOUT_MS) {
  await waitUntil(async () => {
    return await locator.isVisible().catch(() => false) && await locator.isEnabled().catch(() => false);
  }, `تفعيل ${description}`, timeoutMs);
}

async function bodyContains(page, wanted) {
  const body = normalizeText(await page.locator('body').innerText().catch(() => ''));
  const list = Array.isArray(wanted) ? wanted : [wanted];
  return list.some((x) => body.includes(normalizeText(x)));
}

async function safeClickAndWait(page, locator, description, {
  expectText = null,
  expectUrlPart = null,
  expectFn = null,
  timeoutMs = SITE_ACTION_TIMEOUT_MS,
} = {}) {
  await waitUntilEnabled(locator, description, timeoutMs);

  const label = await controlLabel(locator);
  if (isFinalSaveLabel(label)) {
    throw new Error(`حماية الأمان منعت الضغط على "${label}". زر الحفظ النهائي ممنوع تماماً.`);
  }

  log(`⏳ ${description}...`);
  await locator.click({ timeout: timeoutMs });

  if (expectUrlPart) {
    await waitUntil(() => Promise.resolve(page.url().includes(expectUrlPart)), `انتقال الصفحة بعد ${description}`, timeoutMs);
  }

  if (expectText) {
    await waitUntil(() => bodyContains(page, expectText), `ظهور نتيجة ${description}`, timeoutMs);
  }

  if (expectFn) {
    await waitUntil(expectFn, `اكتمال ${description}`, timeoutMs);
  }

  await sleep(AFTER_ACTION_SETTLE_MS);
}

async function findClickableByText(page, wantedText, { exact = true } = {}) {
  const wanted = normalizeText(wantedText);
  const controls = page.locator('button, input[type="submit"], input[type="button"], a');
  const count = await controls.count();

  for (let i = 0; i < count; i++) {
    const el = controls.nth(i);
    try {
      if (!(await el.isVisible())) continue;
      const label = normalizeText(await controlLabel(el));
      if (isFinalSaveLabel(label)) continue;
      if ((exact && label === wanted) || (!exact && label.includes(wanted))) return el;
    } catch (_) {}
  }
  return null;
}

async function login(page, config) {
  log('🌐 فتح صفحة الدخول...');
  await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded', timeout: SITE_ACTION_TIMEOUT_MS });

  const passwordInput = page.locator('input[type="password"]:visible').first();
  await waitUntil(() => passwordInput.isVisible().catch(() => false), 'تحميل حقل كلمة المرور');

  const textInputs = page.locator('input[type="text"]:visible, input:not([type]):visible');
  let usernameInput = null;
  const textCount = await textInputs.count();
  for (let i = 0; i < textCount; i++) {
    const candidate = textInputs.nth(i);
    if (await candidate.isVisible().catch(() => false)) {
      usernameInput = candidate;
      break;
    }
  }
  if (!usernameInput) throw new Error('لم أجد حقل اسم المستخدم في صفحة الدخول.');

  await usernameInput.fill(config.username);
  await passwordInput.fill(config.password);

  const loginButton = await findClickableByText(page, 'دخول', { exact: true });
  if (!loginButton) throw new Error('لم أجد زر "دخول".');

  await safeClickAndWait(page, loginButton, 'تسجيل الدخول', {
    expectText: ['القائمة الرئيسية', 'شاشة المعيل', 'شاشة المستفيدين'],
  });

  const body = normalizeText(await page.locator('body').innerText().catch(() => ''));
  if (body.includes('الدخول للنظام') && !body.includes('القائمة الرئيسية')) {
    throw new Error('بقيت صفحة الدخول ظاهرة. تحقق من اسم المستخدم وكلمة المرور.');
  }
  log('✅ تم تسجيل الدخول.');
}

async function navigateToDisbursement(page) {
  if (page.url().includes(DISBURSEMENT_URL_HINT)) return;

  log('➡️ الدخول إلى شاشة الصرفية...');
  const menuButton = await findClickableByText(page, 'شاشة الصرفية', { exact: true });
  if (!menuButton) throw new Error('لم أجد خيار "شاشة الصرفية" بعد تسجيل الدخول.');

  await safeClickAndWait(page, menuButton, 'فتح شاشة الصرفية', {
    expectText: ['شاشة الصرفيات', 'شاشه الصرفيات'],
  });

  const body = normalizeText(await page.locator('body').innerText().catch(() => ''));
  if (!page.url().includes('Family_sub2.aspx') && !body.includes('شاشه الصرفيات') && !body.includes('شاشة الصرفيات')) {
    throw new Error('تم الضغط على شاشة الصرفية لكن الصفحة المطلوبة لم تظهر.');
  }
  log('✅ تم فتح شاشة الصرفية.');
}

async function describeSelect(select) {
  return select.evaluate((el) => {
    const opts = Array.from(el.options || []).map((o) => (o.textContent || '').trim());
    const parentText = (el.closest('tr')?.innerText || el.parentElement?.innerText || '').trim();
    return { id: el.id || '', name: el.name || '', options: opts, parentText };
  }).catch(() => ({ id: '', name: '', options: [], parentText: '' }));
}

async function findCategorySelect(page) {
  const selects = page.locator('select:visible');
  const count = await selects.count();
  let best = null;
  let bestScore = -1;

  for (let i = 0; i < count; i++) {
    const s = selects.nth(i);
    const d = await describeSelect(s);
    const options = d.options.map(normalizeText);
    let score = 0;
    if (options.some((x) => x.includes('ايتام'))) score += 4;
    if (options.some((x) => x === 'اسر' || x.includes('اسر'))) score += 4;
    if (options.some((x) => x.includes('طلاب علم'))) score += 4;
    if (normalizeText(d.parentText).includes('تصنيف الاسره')) score += 5;
    if (score > bestScore) {
      bestScore = score;
      best = s;
    }
  }
  return bestScore >= 4 ? best : null;
}

async function findMonthSelect(page) {
  const selects = page.locator('select:visible');
  const count = await selects.count();
  let best = null;
  let bestScore = -1;

  for (let i = 0; i < count; i++) {
    const s = selects.nth(i);
    const d = await describeSelect(s);
    const options = d.options.map((x) => normalizeDigits(x)).filter(Boolean);
    const unique = new Set(options);
    let score = 0;
    if ([...Array(12)].every((_, idx) => unique.has(String(idx + 1)))) score += 8;
    if (normalizeText(d.parentText).includes('شهر')) score += 4;
    if (score > bestScore) {
      bestScore = score;
      best = s;
    }
  }
  return bestScore >= 8 ? best : null;
}

async function findYearInput(page) {
  const inputs = page.locator('input:visible');
  const count = await inputs.count();
  let best = null;
  let bestScore = -1;

  for (let i = 0; i < count; i++) {
    const input = inputs.nth(i);
    try {
      const type = (await input.getAttribute('type')) || 'text';
      if (/hidden|submit|button|password|checkbox|radio|file/i.test(type)) continue;
      const value = await input.inputValue().catch(() => '');
      const context = await input.evaluate((el) => (el.closest('tr')?.innerText || el.parentElement?.innerText || '').trim()).catch(() => '');
      let score = 0;
      if (/^\d{4}$/.test(normalizeDigits(value))) score += 5;
      if (normalizeText(context).includes('سنه')) score += 6;
      if (score > bestScore) {
        bestScore = score;
        best = input;
      }
    } catch (_) {}
  }
  return bestScore >= 5 ? best : null;
}

async function selectedOptionMatches(select, wantedText) {
  const wanted = normalizeText(wantedText);
  const selected = normalizeText(await select.locator('option:checked').first().textContent().catch(() => ''));
  return selected === wanted || selected.includes(wanted) || wanted.includes(selected);
}

async function selectOptionFast(page, select, value, wantedText, description) {
  await waitUntilEnabled(select, description);
  await select.selectOption(value);
  await waitUntil(() => selectedOptionMatches(select, wantedText), `تأكيد ${description}`);
  await sleep(FIELD_SETTLE_MS);
}

async function selectByNormalizedText(page, select, wanted, description) {
  const target = normalizeText(wanted);
  const options = await select.locator('option').all();

  for (const option of options) {
    const text = normalizeText(await option.textContent().catch(() => ''));
    if (!text) continue;
    if (text === target || text.includes(target) || target.includes(text)) {
      const value = await option.getAttribute('value');
      if (value !== null) {
        await selectOptionFast(page, select, value, wanted, description);
        return true;
      }
    }
  }
  return false;
}

async function disbursementRowsMatchConfig(page, config) {
  return await page.evaluate(({ year, month, category }) => {
    const norm = (value) => String(value || '')
      .trim()
      .replace(/[أإآ]/g, 'ا')
      .replace(/ى/g, 'ي')
      .replace(/ة/g, 'ه')
      .replace(/ـ/g, '')
      .replace(/\s+/g, ' ');
    const digits = (value) => String(value || '').replace(/[٠-٩]/g, (d) => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[^\d]/g, '');

    const wantedCategory = norm(category);
    const tables = Array.from(document.querySelectorAll('table'));

    for (const table of tables) {
      const rows = Array.from(table.querySelectorAll('tr'));
      let headerCells = null;
      for (const row of rows) {
        const cells = Array.from(row.children);
        const texts = cells.map((c) => norm(c.innerText || c.textContent || ''));
        if (texts.some((t) => t.includes('الرقم الوطني')) && texts.some((t) => t.includes('القيم'))) {
          headerCells = texts;
          break;
        }
      }
      if (!headerCells) continue;

      const idxNat = headerCells.findIndex((t) => t.includes('الرقم الوطني'));
      const idxMonth = headerCells.findIndex((t) => t === 'الشهر' || t.includes('الشهر'));
      const idxYear = headerCells.findIndex((t) => t === 'السنه' || t === 'السنة' || t.includes('السنه') || t.includes('السنة'));
      const idxCategory = headerCells.findIndex((t) => t.includes('التصنيف'));

      for (const row of rows) {
        const cells = Array.from(row.children);
        if (cells.length <= Math.max(idxNat, idxMonth, idxYear, idxCategory)) continue;
        const nat = idxNat >= 0 ? digits(cells[idxNat]?.innerText || cells[idxNat]?.textContent || '') : '';
        if (nat.length < 8) continue;
        const m = idxMonth >= 0 ? digits(cells[idxMonth]?.innerText || cells[idxMonth]?.textContent || '') : '';
        const y = idxYear >= 0 ? digits(cells[idxYear]?.innerText || cells[idxYear]?.textContent || '') : '';
        const c = idxCategory >= 0 ? norm(cells[idxCategory]?.innerText || cells[idxCategory]?.textContent || '') : '';
        if (m === String(month) && y === String(year) && (c === wantedCategory || c.includes(wantedCategory) || wantedCategory.includes(c))) {
          return true;
        }
      }
    }
    return false;
  }, config).catch(() => false);
}

async function configureDisbursement(page, config) {
  log(`⚙️ تجهيز الصرفية: سنة ${config.year}، شهر ${config.month}، تصنيف ${config.category}`);

  const yearInput = await findYearInput(page);
  if (!yearInput) throw new Error('لم أجد حقل السنة في شاشة الصرفية.');
  await waitUntilEnabled(yearInput, 'حقل السنة');
  await yearInput.fill(config.year);

  const monthSelect = await findMonthSelect(page);
  if (!monthSelect) throw new Error('لم أجد قائمة الشهر في شاشة الصرفية.');

  let monthSelected = false;
  const monthOptions = await monthSelect.locator('option').all();
  for (const option of monthOptions) {
    const label = normalizeDigits(await option.textContent().catch(() => ''));
    if (label === config.month) {
      const value = await option.getAttribute('value');
      if (value !== null) {
        await selectOptionFast(page, monthSelect, value, config.month, `اختيار الشهر ${config.month}`);
        monthSelected = true;
        break;
      }
    }
  }
  if (!monthSelected) throw new Error(`لم أجد الشهر ${config.month} ضمن قائمة الأشهر.`);

  const categorySelect = await findCategorySelect(page);
  if (!categorySelect) throw new Error('لم أجد قائمة تصنيف الأسرة.');
  if (!(await selectByNormalizedText(page, categorySelect, config.category, `اختيار التصنيف ${config.category}`))) {
    throw new Error(`لم أجد التصنيف "${config.category}" في قائمة التصنيفات.`);
  }

  const showButton = await findClickableByText(page, 'عرض', { exact: true });
  if (!showButton) throw new Error('لم أجد زر "عرض" في شاشة الصرفية.');

  await safeClickAndWait(page, showButton, 'عرض الصرفية', {
    expectFn: () => disbursementRowsMatchConfig(page, config),
  });

  log('✅ تم عرض الصرفية المطلوبة.');
}

function nameLooksLikeMatch(rowText, expectedName) {
  const expected = normalizeText(expectedName);
  if (!expected) return true;

  const row = normalizeText(rowText);
  if (row.includes(expected)) return true;

  const tokens = expected.split(' ').filter((x) => x.length >= 3);
  if (!tokens.length) return true;
  const matched = tokens.filter((token) => row.includes(token)).length;
  return matched >= Math.min(3, tokens.length);
}

async function findRowByNatId(page, record) {
  const rows = page.locator('tr:visible');
  const count = await rows.count();

  for (let i = 0; i < count; i++) {
    const row = rows.nth(i);
    try {
      const cells = row.locator('td, th');
      const cellCount = await cells.count();
      let natFound = false;

      for (let c = 0; c < cellCount; c++) {
        const text = await cells.nth(c).textContent().catch(() => '');
        if (normalizeDigits(text) === record.natId) {
          natFound = true;
          break;
        }
      }
      if (!natFound) continue;

      const rowText = await row.textContent().catch(() => '');
      if (!nameLooksLikeMatch(rowText, record.name)) {
        throw new Error(`وجدت الرقم الوطني ${record.natId} لكن الاسم في الصف لا يطابق الاسم في ملف المقارنة.`);
      }
      return row;
    } catch (err) {
      if (String(err.message || err).includes('الاسم في الصف')) throw err;
    }
  }
  return null;
}

async function locateAmountInputIndex(row, oldAmount) {
  return row.evaluate((tr, expectedOld) => {
    const normText = (value) => String(value || '')
      .trim()
      .replace(/[أإآ]/g, 'ا')
      .replace(/ى/g, 'ي')
      .replace(/ة/g, 'ه')
      .replace(/ـ/g, '')
      .replace(/\s+/g, ' ');
    const normAmount = (value) => {
      const s = String(value ?? '')
        .trim()
        .replace(/[٠-٩]/g, (d) => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))
        .replace(/,/g, '')
        .replace(/[^\d.-]/g, '');
      if (!s) return '';
      const n = Number(s);
      return Number.isFinite(n) ? String(n) : s;
    };
    const visible = (el) => {
      const r = el.getBoundingClientRect();
      const s = window.getComputedStyle(el);
      return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
    };

    const allInputs = Array.from(tr.querySelectorAll('input'));
    const candidates = allInputs.filter((el) => {
      const type = String(el.getAttribute('type') || 'text').toLowerCase();
      if (/hidden|button|submit|checkbox|radio|image/.test(type)) return false;
      if (el.disabled || el.readOnly || !visible(el)) return false;
      return true;
    });

    if (!candidates.length) return { index: -1, reason: 'no-editable-inputs', debug: [] };

    const debug = candidates.map((el) => ({
      value: normAmount(el.value),
      id: el.id || '',
      name: el.name || '',
      x: Math.round(el.getBoundingClientRect().left + el.getBoundingClientRect().width / 2),
    }));

    if (expectedOld !== '') {
      const matches = candidates.filter((el) => normAmount(el.value) === expectedOld);
      if (matches.length === 1) {
        return { index: allInputs.indexOf(matches[0]), reason: 'old-amount-exact', debug };
      }
    }

    const table = tr.closest('table');
    if (!table) return { index: -1, reason: 'no-table', debug };

    let valueHeader = null;
    let notesHeader = null;
    for (const headerRow of Array.from(table.querySelectorAll('tr'))) {
      const cells = Array.from(headerRow.children);
      const texts = cells.map((c) => normText(c.innerText || c.textContent || ''));
      if (!texts.some((t) => t.includes('الرقم الوطني'))) continue;
      for (let i = 0; i < cells.length; i++) {
        const t = texts[i];
        if (!valueHeader && (t === 'القيمه' || t === 'القيمة' || t.includes('القيمه') || t.includes('القيمة'))) valueHeader = cells[i];
        if (!notesHeader && t.includes('ملاحظ')) notesHeader = cells[i];
      }
      if (valueHeader) break;
    }

    if (!valueHeader) return { index: -1, reason: 'value-header-not-found', debug };

    const valueRect = valueHeader.getBoundingClientRect();
    const valueX = valueRect.left + valueRect.width / 2;
    let notesX = null;
    if (notesHeader) {
      const r = notesHeader.getBoundingClientRect();
      notesX = r.left + r.width / 2;
    }

    const ranked = candidates.map((el) => {
      const r = el.getBoundingClientRect();
      const x = r.left + r.width / 2;
      const valueDistance = Math.abs(x - valueX);
      const notesDistance = notesX === null ? Infinity : Math.abs(x - notesX);
      return { el, valueDistance, notesDistance };
    }).sort((a, b) => a.valueDistance - b.valueDistance);

    const best = ranked[0];
    if (!best) return { index: -1, reason: 'no-ranked-input', debug };

    if (best.valueDistance > 180 || best.notesDistance < best.valueDistance) {
      return { index: -1, reason: 'geometry-not-confident', debug };
    }

    return { index: allInputs.indexOf(best.el), reason: 'value-column-geometry', debug };
  }, oldAmount).catch(() => ({ index: -1, reason: 'evaluate-failed', debug: [] }));
}

async function findAmountInput(row, record) {
  const located = await locateAmountInputIndex(row, record.oldAmount);
  if (located.index >= 0) {
    const input = row.locator('input').nth(located.index);
    if (await input.isEditable().catch(() => false)) {
      log(`🔎 حقل القيمة ${record.natId}: ${located.reason}`);
      return input;
    }
  }

  const details = (located.debug || []).map((x) => `${x.value || '(فارغ)'}[${x.id || x.name || 'بدون اسم'}]`).join('، ');
  throw new Error(
    `لم أستطع تحديد حقل القيمة بثقة للرقم الوطني ${record.natId}.` +
    (details ? ` الحقول الموجودة في الصف: ${details}` : '')
  );
}

async function processRecord(page, record, index, total) {
  progress(index, total, record, 'جاري البحث');

  const row = await findRowByNatId(page, record);
  if (!row) throw new Error(`لم أجد الرقم الوطني ${record.natId} داخل الصرفية المعروضة.`);

  const input = await findAmountInput(row, record);
  await waitUntil(() => input.isEditable().catch(() => false), `جاهزية حقل قيمة الرقم ${record.natId}`);

  const current = normalizeAmount(await input.inputValue().catch(() => ''));
  const target = normalizeAmount(record.newAmount);

  if (current === target) {
    progress(index, total, record, 'صحيح مسبقاً');
    return { ...record, status: 'already_correct', current };
  }

  if (record.oldAmount !== '' && current !== record.oldAmount) {
    throw new Error(`القيمة الحالية في الحقل المكتشف (${current}) لا تطابق قيمة ملف كرامة (${record.oldAmount}) للرقم ${record.natId}.`);
  }

  await input.fill(target);
  await waitUntil(async () => normalizeAmount(await input.inputValue().catch(() => '')) === target, `تأكيد تعديل الرقم ${record.natId}`, 5000);
  await sleep(FIELD_SETTLE_MS);

  const after = normalizeAmount(await input.inputValue().catch(() => ''));
  if (after !== target) {
    throw new Error(`تمت محاولة تعديل ${record.natId} إلى ${target} لكن الحقل أصبح ${after}.`);
  }

  progress(index, total, record, 'تم التعديل');
  return { ...record, status: 'changed', current, newAmount: target };
}

async function clickTemporarySave(page) {
  log('💾 انتهت التعديلات. جاري الضغط على "حفظ مؤقت" فقط...');
  const controls = page.locator('button, input[type="submit"], input[type="button"]');
  const count = await controls.count();
  let tempButton = null;

  for (let i = 0; i < count; i++) {
    const el = controls.nth(i);
    try {
      if (!(await el.isVisible())) continue;
      const label = normalizeText(await controlLabel(el));
      if (isFinalSaveLabel(label)) continue;
      if (label === 'حفظ مؤقت' || (label.includes('حفظ') && label.includes('مؤقت'))) {
        tempButton = el;
        break;
      }
    } catch (_) {}
  }

  if (!tempButton) throw new Error('لم أجد زر "حفظ مؤقت". لم يتم تنفيذ أي حفظ نهائي.');
  const chosenLabel = normalizeText(await controlLabel(tempButton));
  if (isFinalSaveLabel(chosenLabel)) {
    throw new Error('حماية الأمان منعت الضغط لأن الزر المكتشف يحتوي على كلمة "نهائي".');
  }

  const successText = 'تم حفظ القيم بنجاح';
  const successWasVisibleBefore = await bodyContains(page, successText);
  let navigated = false;
  const navPromise = page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: SITE_ACTION_TIMEOUT_MS })
    .then(() => { navigated = true; return true; })
    .catch(() => false);

  await waitUntilEnabled(tempButton, 'زر حفظ مؤقت');
  log('⏳ الحفظ المؤقت...');
  await tempButton.click({ timeout: SITE_ACTION_TIMEOUT_MS });

  await waitUntil(async () => {
    if (await bodyContains(page, successText)) {
      if (!successWasVisibleBefore || navigated) return true;
    }
    return false;
  }, 'رسالة نجاح الحفظ المؤقت');

  await Promise.race([navPromise, sleep(200)]).catch(() => {});
  if (!(await bodyContains(page, successText))) {
    throw new Error('تم الضغط على حفظ مؤقت، لكن لم تظهر رسالة "تم حفظ القيم بنجاح".');
  }

  log('✅ تم الحفظ المؤقت وظهرت رسالة نجاح الحفظ.');
}

async function readVisibleTotal(page) {
  try {
    return await page.evaluate(() => {
      const norm = (v) => String(v || '').trim().replace(/\s+/g, ' ');
      const nodes = Array.from(document.querySelectorAll('td,th,span,label,div'));
      const label = nodes.find((el) => norm(el.innerText || el.textContent || '').includes('المجموع الكلي للصرفية'));
      if (!label) return '';

      const row = label.closest('tr') || label.parentElement;
      if (!row) return '';
      const input = row.querySelector('input');
      if (input && input.value) return input.value;

      const text = norm(row.innerText || row.textContent || '');
      const match = text.match(/(?:المجموع الكلي للصرفية)\s*([\d.,]+)/);
      return match ? match[1] : '';
    });
  } catch (_) {
    return '';
  }
}

async function main() {
  const config = readConfig();
  const records = loadRecords(config.excelPath);
  if (!records.length) throw new Error('لا توجد سجلات في ورقة "يحتاج تعديل".');

  log(`📄 تم تحميل ${records.length} سجل يحتاج تعديل.`);

  const duplicateIds = records.map((r) => r.natId).filter((id, idx, arr) => arr.indexOf(id) !== idx);
  if (duplicateIds.length) {
    throw new Error(`يوجد رقم وطني مكرر في ملف التعديل: ${[...new Set(duplicateIds)].join(', ')}. تم إيقاف العملية للحماية.`);
  }

  const browser = await launchBrowser();
  const page = await browser.newPage();
  page.setDefaultTimeout(SITE_ACTION_TIMEOUT_MS);
  page.setDefaultNavigationTimeout(SITE_ACTION_TIMEOUT_MS);

  const results = [];
  let temporarySaved = false;

  try {
    await login(page, config);
    await navigateToDisbursement(page);
    await configureDisbursement(page, config);

    for (let i = 0; i < records.length; i++) {
      const record = records[i];
      try {
        const result = await processRecord(page, record, i + 1, records.length);
        results.push(result);
      } catch (err) {
        results.push({ ...record, status: 'error', error: err.message || String(err) });
        const reportPath = outputPath(`auto_update_result_${Date.now()}.json`);
        fs.writeFileSync(reportPath, JSON.stringify(results, null, 2), 'utf8');
        throw new Error(
          `توقفت العملية عند السجل ${i + 1}/${records.length}: ${err.message || err}\n` +
          'لم يتم الضغط على حفظ مؤقت بعد هذا الخطأ.'
        );
      }
    }

    await clickTemporarySave(page);
    temporarySaved = true;

    const total = await readVisibleTotal(page);
    const reportPath = outputPath(`auto_update_result_${Date.now()}.json`);
    fs.writeFileSync(reportPath, JSON.stringify({
      year: config.year,
      month: config.month,
      category: config.category,
      temporarySaved: true,
      visibleTotal: total,
      results,
    }, null, 2), 'utf8');

    console.log(`__TEMP_SAVE_OK__|${total || ''}|${reportPath}`);
    log('🛑 انتهت مسؤولية البرنامج هنا. لن يتم الضغط على "حفظ نهائي" بأي شكل.');
    log('👀 المتصفح سيبقى مفتوحاً لتراجع المجموع ثم تقوم أنت بالحفظ النهائي يدوياً إذا كان صحيحاً.');

    while (browser.isConnected()) await sleep(1000);
  } catch (err) {
    const reportPath = outputPath(`auto_update_error_${Date.now()}.json`);
    fs.writeFileSync(reportPath, JSON.stringify({
      year: config.year,
      month: config.month,
      category: config.category,
      temporarySaved,
      error: err.message || String(err),
      results,
    }, null, 2), 'utf8');
    throw err;
  }
}

main().catch((err) => {
  const message = String(err.message || err).replace(/[\r\n]+/g, ' ');
  console.error(`__FATAL__|${message}`);
  console.error(`❌ ${message}`);
  process.exitCode = 1;
});
