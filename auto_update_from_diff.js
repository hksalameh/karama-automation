const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const xlsx = require('xlsx');

const LOGIN_URL = 'http://92.253.101.105:8040/ICCS/index.aspx';
const DISBURSEMENT_URL_HINT = '/ICCS/Family_sub2.aspx';

// هذه حدود قصوى للحماية فقط؛ لا يوجد انتظار ثابت بين الخطوات.
const SITE_ACTION_TIMEOUT_MS = 20000;
const FIELD_CONFIRM_TIMEOUT_MS = 3000;
const POLL_MS = 100;

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
      if (await check()) return;
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

function detectHeader(headers, predicates) {
  for (const header of headers) {
    const normalized = normalizeText(header);
    if (predicates.some((predicate) => predicate(normalized))) return header;
  }
  return null;
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
    .filter((record) => record.natId && record.newAmount !== '');
}

async function launchBrowser() {
  const attempts = [
    ['Chrome', { channel: 'chrome', headless: false }],
    ['Edge', { channel: 'msedge', headless: false }],
    ['Chromium', { headless: false }],
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

async function findClickableByText(page, wantedText, { exact = true } = {}) {
  const wanted = normalizeText(wantedText);
  const controls = page.locator('button, input[type="submit"], input[type="button"], a');
  const count = await controls.count();

  for (let i = 0; i < count; i++) {
    const control = controls.nth(i);
    try {
      if (!(await control.isVisible())) continue;
      if (!(await control.isEnabled())) continue;
      const label = normalizeText(await controlLabel(control));
      if (isFinalSaveLabel(label)) continue;
      if ((exact && label === wanted) || (!exact && label.includes(wanted))) return control;
    } catch (_) {}
  }
  return null;
}

async function bodyContains(page, wanted) {
  const body = normalizeText(await page.locator('body').innerText().catch(() => ''));
  const values = Array.isArray(wanted) ? wanted : [wanted];
  return values.some((value) => body.includes(normalizeText(value)));
}

async function login(page, config) {
  log('🌐 فتح صفحة الدخول...');
  await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded', timeout: SITE_ACTION_TIMEOUT_MS });

  const passwordInput = page.locator('input[type="password"]:visible').first();
  await waitUntil(() => passwordInput.isVisible().catch(() => false), 'ظهور حقل كلمة المرور');

  const textInputs = page.locator('input[type="text"]:visible, input:not([type]):visible');
  let usernameInput = null;
  for (let i = 0; i < await textInputs.count(); i++) {
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

  log('⏳ تسجيل الدخول...');
  await loginButton.click({ timeout: SITE_ACTION_TIMEOUT_MS });
  await waitUntil(
    () => bodyContains(page, ['القائمة الرئيسية', 'شاشة المعيل', 'شاشة المستفيدين']),
    'ظهور الصفحة الرئيسية بعد تسجيل الدخول'
  );
  log('✅ تم تسجيل الدخول.');
}

async function navigateToDisbursement(page) {
  if (page.url().includes(DISBURSEMENT_URL_HINT)) return;

  log('➡️ الدخول إلى شاشة الصرفية...');
  const menuButton = await findClickableByText(page, 'شاشة الصرفية', { exact: true });
  if (!menuButton) throw new Error('لم أجد خيار "شاشة الصرفية" بعد تسجيل الدخول.');

  await menuButton.click({ timeout: SITE_ACTION_TIMEOUT_MS });
  await waitUntil(async () => {
    if (page.url().includes('Family_sub2.aspx')) return true;
    return bodyContains(page, ['شاشة الصرفيات', 'شاشه الصرفيات']);
  }, 'ظهور شاشة الصرفية');

  log('✅ تم فتح شاشة الصرفية.');
}

async function describeSelect(select) {
  return select.evaluate((el) => ({
    options: Array.from(el.options || []).map((option) => (option.textContent || '').trim()),
    parentText: (el.closest('tr')?.innerText || el.parentElement?.innerText || '').trim(),
  })).catch(() => ({ options: [], parentText: '' }));
}

async function findCategorySelect(page) {
  const selects = page.locator('select:visible');
  let best = null;
  let bestScore = -1;

  for (let i = 0; i < await selects.count(); i++) {
    const select = selects.nth(i);
    const info = await describeSelect(select);
    const options = info.options.map(normalizeText);
    let score = 0;
    if (options.some((x) => x.includes('ايتام'))) score += 4;
    if (options.some((x) => x === 'اسر' || x.includes('اسر'))) score += 4;
    if (options.some((x) => x.includes('طلاب علم'))) score += 4;
    if (normalizeText(info.parentText).includes('تصنيف الاسره')) score += 5;
    if (score > bestScore) {
      bestScore = score;
      best = select;
    }
  }
  return bestScore >= 4 ? best : null;
}

async function findMonthSelect(page) {
  const selects = page.locator('select:visible');
  let best = null;
  let bestScore = -1;

  for (let i = 0; i < await selects.count(); i++) {
    const select = selects.nth(i);
    const info = await describeSelect(select);
    const values = new Set(info.options.map((x) => normalizeDigits(x)).filter(Boolean));
    let score = 0;
    if ([...Array(12)].every((_, idx) => values.has(String(idx + 1)))) score += 8;
    if (normalizeText(info.parentText).includes('شهر')) score += 4;
    if (score > bestScore) {
      bestScore = score;
      best = select;
    }
  }
  return bestScore >= 8 ? best : null;
}

async function findYearInput(page) {
  const inputs = page.locator('input:visible');
  let best = null;
  let bestScore = -1;

  for (let i = 0; i < await inputs.count(); i++) {
    const input = inputs.nth(i);
    try {
      const type = ((await input.getAttribute('type')) || 'text').toLowerCase();
      if (/hidden|submit|button|password|checkbox|radio|file/.test(type)) continue;
      const current = normalizeDigits(await input.inputValue().catch(() => ''));
      const context = normalizeText(await input.evaluate((el) => (el.closest('tr')?.innerText || el.parentElement?.innerText || '')).catch(() => ''));
      let score = 0;
      if (/^\d{4}$/.test(current)) score += 5;
      if (context.includes('سنه') || context.includes('سنة')) score += 6;
      if (score > bestScore) {
        bestScore = score;
        best = input;
      }
    } catch (_) {}
  }
  return bestScore >= 5 ? best : null;
}

async function selectMonth(select, month) {
  const options = await select.locator('option').all();
  for (const option of options) {
    if (normalizeDigits(await option.textContent().catch(() => '')) !== month) continue;
    const value = await option.getAttribute('value');
    if (value === null) continue;
    await select.selectOption(value);
    await waitUntil(
      async () => normalizeDigits(await select.locator('option:checked').first().textContent().catch(() => '')) === month,
      `تأكيد اختيار الشهر ${month}`,
      FIELD_CONFIRM_TIMEOUT_MS
    );
    return true;
  }
  return false;
}

async function selectCategory(select, category) {
  const wanted = normalizeText(category);
  const options = await select.locator('option').all();
  for (const option of options) {
    const text = normalizeText(await option.textContent().catch(() => ''));
    if (!(text === wanted || text.includes(wanted) || wanted.includes(text))) continue;
    const value = await option.getAttribute('value');
    if (value === null) continue;
    await select.selectOption(value);
    await waitUntil(async () => {
      const selected = normalizeText(await select.locator('option:checked').first().textContent().catch(() => ''));
      return selected === wanted || selected.includes(wanted) || wanted.includes(selected);
    }, `تأكيد اختيار التصنيف ${category}`, FIELD_CONFIRM_TIMEOUT_MS);
    return true;
  }
  return false;
}

async function indexDisbursementTable(page, config) {
  return page.evaluate(({ year, month, category }) => {
    const normText = (value) => String(value || '')
      .trim()
      .replace(/[أإآ]/g, 'ا')
      .replace(/ى/g, 'ي')
      .replace(/ة/g, 'ه')
      .replace(/ـ/g, '')
      .replace(/\s+/g, ' ');
    const digits = (value) => String(value || '')
      .replace(/[٠-٩]/g, (d) => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))
      .replace(/[^\d]/g, '');
    const wantedCategory = normText(category);

    for (const old of document.querySelectorAll('[data-karama-nat]')) old.removeAttribute('data-karama-nat');
    for (const old of document.querySelectorAll('[data-karama-amount]')) old.removeAttribute('data-karama-amount');

    for (const table of Array.from(document.querySelectorAll('table'))) {
      const rows = Array.from(table.rows || []);
      let header = null;
      let headerRow = null;

      for (const row of rows) {
        const cells = Array.from(row.cells || []);
        const texts = cells.map((cell) => normText(cell.innerText || cell.textContent || ''));
        if (texts.some((x) => x.includes('الرقم الوطني')) && texts.some((x) => x.includes('القيم'))) {
          header = texts;
          headerRow = row;
          break;
        }
      }
      if (!header || !headerRow) continue;

      const idxNat = header.findIndex((x) => x.includes('الرقم الوطني'));
      const idxValue = header.findIndex((x) => x === 'القيمه' || x === 'القيمة' || x.includes('القيمه') || x.includes('القيمة'));
      const idxMonth = header.findIndex((x) => x.includes('الشهر'));
      const idxYear = header.findIndex((x) => x.includes('السنه') || x.includes('السنة'));
      const idxCategory = header.findIndex((x) => x.includes('التصنيف'));

      if (idxNat < 0 || idxValue < 0) continue;

      const seen = new Map();
      let matchedConfig = false;
      let indexedCount = 0;

      for (const row of rows) {
        if (row === headerRow) continue;
        const cells = Array.from(row.cells || []);
        if (cells.length <= Math.max(idxNat, idxValue)) continue;

        const nat = digits(cells[idxNat]?.innerText || cells[idxNat]?.textContent || '');
        if (nat.length < 8) continue;

        const rowMonth = idxMonth >= 0 && idxMonth < cells.length ? digits(cells[idxMonth]?.innerText || cells[idxMonth]?.textContent || '') : '';
        const rowYear = idxYear >= 0 && idxYear < cells.length ? digits(cells[idxYear]?.innerText || cells[idxYear]?.textContent || '') : '';
        const rowCategory = idxCategory >= 0 && idxCategory < cells.length ? normText(cells[idxCategory]?.innerText || cells[idxCategory]?.textContent || '') : '';

        if ((idxMonth < 0 || rowMonth === String(month)) &&
            (idxYear < 0 || rowYear === String(year)) &&
            (idxCategory < 0 || rowCategory === wantedCategory || rowCategory.includes(wantedCategory) || wantedCategory.includes(rowCategory))) {
          matchedConfig = true;
        }

        seen.set(nat, (seen.get(nat) || 0) + 1);
        row.setAttribute('data-karama-nat', nat);

        const valueCell = cells[idxValue];
        const candidates = Array.from(valueCell?.querySelectorAll('input') || []).filter((input) => {
          const type = String(input.getAttribute('type') || 'text').toLowerCase();
          return !/hidden|button|submit|checkbox|radio|image/.test(type) && !input.disabled && !input.readOnly;
        });
        if (candidates.length === 1) {
          candidates[0].setAttribute('data-karama-amount', '1');
        }
        indexedCount += 1;
      }

      const duplicates = Array.from(seen.entries()).filter(([, count]) => count > 1).map(([nat]) => nat);
      return { found: true, matchedConfig, indexedCount, duplicates };
    }

    return { found: false, matchedConfig: false, indexedCount: 0, duplicates: [] };
  }, config).catch(() => ({ found: false, matchedConfig: false, indexedCount: 0, duplicates: [] }));
}

async function configureDisbursement(page, config) {
  log(`⚙️ تجهيز الصرفية: سنة ${config.year}، شهر ${config.month}، تصنيف ${config.category}`);

  const yearInput = await findYearInput(page);
  if (!yearInput) throw new Error('لم أجد حقل السنة في شاشة الصرفية.');
  await yearInput.fill(config.year);

  const monthSelect = await findMonthSelect(page);
  if (!monthSelect || !(await selectMonth(monthSelect, config.month))) {
    throw new Error(`لم أجد الشهر ${config.month} ضمن قائمة الأشهر.`);
  }

  const categorySelect = await findCategorySelect(page);
  if (!categorySelect || !(await selectCategory(categorySelect, config.category))) {
    throw new Error(`لم أجد التصنيف "${config.category}" في قائمة التصنيفات.`);
  }

  const showButton = await findClickableByText(page, 'عرض', { exact: true });
  if (!showButton) throw new Error('لم أجد زر "عرض" في شاشة الصرفية.');

  log('⏳ عرض الصرفية...');
  await showButton.click({ timeout: SITE_ACTION_TIMEOUT_MS });

  let tableInfo = null;
  await waitUntil(async () => {
    tableInfo = await indexDisbursementTable(page, config);
    return tableInfo.found && tableInfo.matchedConfig && tableInfo.indexedCount > 0;
  }, 'ظهور أسماء وبيانات الصرفية المطلوبة');

  if (tableInfo.duplicates.length) {
    throw new Error(`يوجد رقم وطني مكرر في جدول موقع كرامة: ${tableInfo.duplicates.slice(0, 10).join(', ')}`);
  }

  log(`✅ ظهرت الصرفية وتم تجهيز ${tableInfo.indexedCount} سجل للبحث السريع.`);
}

async function processRecordFast(page, record) {
  return page.evaluate((record) => {
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

    const rows = Array.from(document.querySelectorAll(`tr[data-karama-nat="${record.natId}"]`));
    if (rows.length === 0) return { ok: false, error: `لم أجد الرقم الوطني ${record.natId} داخل الصرفية المعروضة.` };
    if (rows.length > 1) return { ok: false, error: `وجدت أكثر من صف للرقم الوطني ${record.natId}.` };

    const row = rows[0];
    const rowText = normText(row.innerText || row.textContent || '');
    const expectedName = normText(record.name);
    if (expectedName) {
      const tokens = expectedName.split(' ').filter((token) => token.length >= 3);
      const exact = rowText.includes(expectedName);
      const matched = tokens.filter((token) => rowText.includes(token)).length;
      if (!exact && tokens.length && matched < Math.min(3, tokens.length)) {
        return { ok: false, error: `وجدت الرقم الوطني ${record.natId} لكن الاسم في نفس الصف لا يطابق ملف المقارنة.` };
      }
    }

    let input = row.querySelector('input[data-karama-amount="1"]');
    if (!input) {
      const candidates = Array.from(row.querySelectorAll('input')).filter((el) => {
        const type = String(el.getAttribute('type') || 'text').toLowerCase();
        return !/hidden|button|submit|checkbox|radio|image/.test(type) && !el.disabled && !el.readOnly;
      });
      const oldMatches = candidates.filter((el) => normAmount(el.value) === record.oldAmount);
      if (oldMatches.length === 1) input = oldMatches[0];
    }

    if (!input) {
      return { ok: false, error: `لم أستطع تحديد حقل القيمة بثقة للرقم الوطني ${record.natId}.` };
    }

    const current = normAmount(input.value);
    const target = normAmount(record.newAmount);
    const old = normAmount(record.oldAmount);

    if (current === target) {
      return { ok: true, status: 'already_correct', current, target };
    }
    if (old !== '' && current !== old) {
      return {
        ok: false,
        error: `القيمة الحالية في صف الرقم ${record.natId} هي (${current}) بينما ملف كرامة يحتوي (${old}). لم يتم التعديل.`
      };
    }

    const descriptor = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
    if (descriptor && descriptor.set) descriptor.set.call(input, target);
    else input.value = target;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));

    const after = normAmount(input.value);
    if (after !== target) {
      return { ok: false, error: `تعذر تثبيت القيمة ${target} للرقم الوطني ${record.natId}.` };
    }

    return { ok: true, status: 'changed', current, target };
  }, record);
}

async function processRecord(page, record, index, total) {
  progress(index, total, record, 'جاري البحث');
  const result = await processRecordFast(page, record);
  if (!result.ok) throw new Error(result.error || `خطأ غير معروف عند الرقم ${record.natId}`);

  if (result.status === 'already_correct') {
    progress(index, total, record, 'صحيح مسبقاً');
    return { ...record, status: 'already_correct', current: result.current };
  }

  progress(index, total, record, 'تم التعديل');
  return { ...record, status: 'changed', current: result.current, newAmount: result.target };
}

async function pagePostbackToken(page) {
  return page.evaluate(() => {
    const valueOf = (name) => document.querySelector(`input[name="${name}"]`)?.value || '';
    const viewState = valueOf('__VIEWSTATE');
    const eventValidation = valueOf('__EVENTVALIDATION');
    return `${location.href}|${viewState.slice(-150)}|${eventValidation.slice(-150)}`;
  }).catch(() => '');
}

async function clickTemporarySave(page) {
  log('💾 انتهت جميع التعديلات. جاري تنفيذ حفظ مؤقت فقط...');
  const tempButton = await findClickableByText(page, 'حفظ مؤقت', { exact: true });
  if (!tempButton) throw new Error('لم أجد زر "حفظ مؤقت". لم يتم تنفيذ أي حفظ نهائي.');

  const label = normalizeText(await controlLabel(tempButton));
  if (isFinalSaveLabel(label)) {
    throw new Error('حماية الأمان منعت الضغط على زر مشتبه أنه حفظ نهائي.');
  }

  const successText = 'تم حفظ القيم بنجاح';
  const beforeToken = await pagePostbackToken(page);
  const successBefore = await bodyContains(page, successText);

  await tempButton.click({ timeout: SITE_ACTION_TIMEOUT_MS });

  await waitUntil(async () => {
    const successNow = await bodyContains(page, successText);
    if (successNow && !successBefore) return true;
    const afterToken = await pagePostbackToken(page);
    return afterToken && beforeToken && afterToken !== beforeToken;
  }, 'استجابة الحفظ المؤقت');

  await waitUntil(() => bodyContains(page, successText), 'رسالة نجاح الحفظ المؤقت');
  log('✅ تم الحفظ المؤقت وظهرت رسالة نجاح الحفظ.');
}

async function readVisibleTotal(page) {
  return page.evaluate(() => {
    const norm = (value) => String(value || '').trim().replace(/\s+/g, ' ');
    const nodes = Array.from(document.querySelectorAll('td,th,span,label,div'));
    const label = nodes.find((el) => norm(el.innerText || el.textContent || '').includes('المجموع الكلي للصرفية'));
    if (!label) return '';
    const row = label.closest('tr') || label.parentElement;
    if (!row) return '';
    const input = row.querySelector('input');
    if (input?.value) return input.value;
    const text = norm(row.innerText || row.textContent || '');
    const match = text.match(/(?:المجموع الكلي للصرفية)\s*([\d.,]+)/);
    return match ? match[1] : '';
  }).catch(() => '');
}

async function main() {
  const config = readConfig();
  const records = loadRecords(config.excelPath);
  if (!records.length) throw new Error('لا توجد سجلات في ورقة "يحتاج تعديل".');

  log(`📄 تم تحميل ${records.length} سجل يحتاج تعديل.`);

  const duplicateIds = records.map((record) => record.natId).filter((id, idx, all) => all.indexOf(id) !== idx);
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

    log('⚡ تم تجهيز الجدول. ستتم التعديلات التالية مباشرة دون انتظار بين سجل وآخر.');

    for (let i = 0; i < records.length; i++) {
      const record = records[i];
      try {
        results.push(await processRecord(page, record, i + 1, records.length));
      } catch (err) {
        results.push({ ...record, status: 'error', error: err.message || String(err) });
        const reportPath = outputPath(`auto_update_result_${Date.now()}.json`);
        fs.writeFileSync(reportPath, JSON.stringify(results, null, 2), 'utf8');
        throw new Error(`توقفت العملية عند السجل ${i + 1}/${records.length}: ${err.message || err}\nلم يتم الضغط على حفظ مؤقت بعد هذا الخطأ.`);
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
