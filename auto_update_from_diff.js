const fs = require('fs');
const readline = require('readline');
const { chromium } = require('playwright');
const xlsx = require('xlsx');
const { exec } = require('child_process');


const args = process.argv.slice(2);

// Read Arabic file path from temp file to avoid encoding issues
function readPathFromFile() {
  try {
    const pathFile = process.env.KARAMA_PATH_FILE;
    if (pathFile) {
      const fs = require('fs');
      const p = fs.readFileSync(pathFile, 'utf-8').trim();
      if (p && fs.existsSync(p)) return p;
    }
  } catch (_) {}
  return args.find((arg) => !arg.startsWith('--')) || '';
}

const CONFIG = {
  excelPath: readPathFromFile(),
  siteUrl: 'http://92.253.101.105:8040/ICCS/Family_sub2.aspx',
  sheetName: 'يحتاج تعديل',
  saveChanges: args.includes('--save') && !args.includes('--no-save'),
  temporarySaveEvery: Number(args.find((arg) => arg.startsWith('--save-every='))?.split('=')[1] || 15),
  confirmEach: args.includes('--confirm-each'),
  requireConfirmBeforeEdit: !args.includes('--yes'),
  allowZeroUpdates: args.includes('--allow-zero'),
  fromGUI: args.includes('--from-gui'),
  timeoutMs: 15000,
};

function outputPath(fileName) {
  const dir = process.env.KARAMA_OUTPUT_DIR || process.cwd();
  try {
    fs.mkdirSync(dir, { recursive: true });
  } catch (_) {}
  return require('path').join(dir, fileName);
}

function ask(question) {
  if (CONFIG.fromGUI) {
    const command = waitForGUICommandObject();
    // We only expect 'approve' or 'skip' here during a confirmation step
    if (command.action === 'approve') return 'yes';
    return 'skip';
  }
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => rl.question(question, (ans) => {
    rl.close();
    resolve(ans.trim());
  }));
}

async function waitForGUICommandObject() {
  const controlFile = process.env.KARAMA_CONTROL_FILE;
  if (!controlFile) {
    console.log('⚠️  لم يتم تحديد ملف التحكم');
    return { action: 'shutdown' };
  }
  
  // Wait indefinitely until GUI writes a command
  while (true) {
    try {
      if (fs.existsSync(controlFile)) {
        const content = fs.readFileSync(controlFile, 'utf-8');
        const command = JSON.parse(content);
        
        console.log(`📨 تم استقبال أمر من واجهة المستخدم: ${command.action}`);
        
        fs.unlinkSync(controlFile);
        return command; // Return the full command object
      }
    } catch (e) {
      // File might be being written, wait and retry
    }
    await new Promise(resolve => setTimeout(resolve, 250));
  }
}

function writeCurrentStatus(record, status) {
  const content = [
    `status: ${status}`,
    `natId: ${record.natId}`,
    `name: ${record.name}`,
    `oldAmount: ${record.oldAmount}`,
    `newAmount: ${record.newAmount}`,
    `time: ${new Date().toLocaleString()}`,
  ].join('\n');
  fs.writeFileSync(outputPath('auto_update_current.txt'), content, 'utf8');
}

function normalizeText(value) {
  return String(value ?? '')
    .trim()
    .replace(/[أإآ]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ة/g, 'ه')
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

function nameLooksLikeMatch(rowText, expectedName) {
  const row = normalizeText(rowText);
  const expected = normalizeText(expectedName);
  if (!expected) return true;
  if (row.includes(expected)) return true;

  const tokens = expected.split(' ').filter((x) => x.length >= 3);
  if (tokens.length === 0) return true;

  const matched = tokens.filter((token) => row.includes(token)).length;
  return matched >= Math.min(3, tokens.length);
}



async function visibleEditableInputs(locator) {
  const inputs = await locator.locator('input').all();
  const out = [];

  for (const input of inputs) {
    try {
      const type = (await input.getAttribute('type')) || 'text';
      if (/hidden|button|submit|checkbox|radio|image/i.test(type)) continue;
      if (!(await input.isVisible())) continue;
      if (!(await input.isEnabled())) continue;
      out.push(input);
    } catch (_) {
      // Ignore detached inputs while the ASP.NET page refreshes.
    }
  }

  return out;
}

function amountFieldScore(hints) {
  const haystack = [
    hints.type,
    hints.name,
    hints.id,
    hints.placeholder,
    hints.ariaLabel,
    hints.title,
    hints.autocomplete,
    hints.labelText,
    hints.headerText,
    hints.cellText,
    hints.value,
  ].filter(Boolean).join(' ');

  const normalized = normalizeText(haystack);
  const headerNormalized = normalizeText(hints.headerText || '');
  const patterns = [
    /amount/i,
    /value/i,
    /kaf/i,
    /مبلغ/i,
    /قيم/i,
    /كفال/i,
    /استحقاق/i,
    /بدل/i,
  ];

  let score = 0;
  for (const pattern of patterns) {
    if (pattern.test(haystack) || pattern.test(normalized)) score += 2;
  }
  if (/مبلغ|قيم|قيمة|amount|value|kaf|كفال/i.test(hints.headerText || '') || /مبلغ|قيم|قيمة|كفال/.test(headerNormalized)) score += 8;
  if (/ملاحظ|note|comment/i.test(hints.headerText || '') || /ملاحظ/.test(headerNormalized)) score -= 10;
  if (/سنة|year|شهر|month|رقم وطني|national|nat|اسم|name|الاسم/i.test(hints.headerText || '') || /سنة|شهر|رقم وطني|اسم/.test(headerNormalized)) score -= 4;
  if (/ملاحظ|ملاحظة|note|comment/i.test(haystack) || /ملاحظ|ملاحظة/.test(normalized)) score -= 3;
  if (/text|number|tel|decimal|currency/i.test(hints.type || '')) score += 1;
  if ((hints.value || '').trim() !== '') score += 1;
  return score;
}

async function describeInputField(input) {
  return input.evaluate((el) => {
    const get = (name) => String(el.getAttribute(name) || '').trim();
    const labels = el.labels ? Array.from(el.labels).map((label) => (label.innerText || label.textContent || '').trim()).filter(Boolean).join(' ') : '';
    const ariaLabel = get('aria-label');
    const textLabel = labels || (el.closest('label')?.innerText || el.closest('label')?.textContent || '').trim();
    const cell = el.closest('td,th');
    const row = el.closest('tr');
    const table = row?.closest('table');
    const cellIndex = cell ? Array.from(cell.parentElement?.children || []).indexOf(cell) : -1;
    let headerText = '';
    if (table && cellIndex >= 0) {
      const headerRows = Array.from(table.querySelectorAll('tr')).filter((tr) => tr.querySelector('th'));
      for (const headerRow of headerRows) {
        const headers = Array.from(headerRow.children);
        if (cellIndex < headers.length) {
          headerText = String(headers[cellIndex]?.innerText || headers[cellIndex]?.textContent || '').trim();
          if (headerText) break;
        }
      }
    }
    return {
      type: get('type'),
      name: get('name'),
      id: get('id'),
      placeholder: get('placeholder'),
      ariaLabel,
      title: get('title'),
      autocomplete: get('autocomplete'),
      labelText: textLabel,
      headerText,
      cellIndex,
      value: String(el.value ?? '').trim(),
      cellText: String(cell?.innerText || cell?.textContent || '').trim(),
      rowText: String(row?.innerText || row?.textContent || '').trim(),
    };
  }).catch(() => ({
    type: '',
    name: '',
    id: '',
    placeholder: '',
    ariaLabel: '',
    title: '',
    autocomplete: '',
    labelText: '',
    headerText: '',
    cellIndex: -1,
    value: '',
    cellText: '',
    rowText: '',
  }));
}

async function findSearchInput(page) {
  const candidates = [
    { name: 'XPath 1 (الرقم الوطني)', locator: page.locator('xpath=//*[contains(normalize-space(.), "الرقم الوطني")]/following::input[not(@type="hidden")][1]') },
    { name: 'XPath 2 (رقم وطني)', locator: page.locator('xpath=//*[contains(normalize-space(.), "رقم وطني")]/following::input[not(@type="hidden")][1]') },
    { name: 'CSS - أول input نصي مرئي', locator: page.locator('input[type="text"]:visible').first() },
    { name: 'CSS - أي input مرئي', locator: page.locator('input:visible').first() },
  ];

  for (const candidate of candidates) {
    try {
      const count = await candidate.locator.count().catch(() => 0);
      if (count > 0 && await candidate.locator.first().isVisible()) {
        console.log(`✅ وجدت حقل البحث: ${candidate.name}`);
        return candidate.locator.first();
      }
    } catch (_) {
      // Try the next locator.
    }
  }

  console.log('⚠️  لم أتمكن من العثور على أي حقل بحث في الصفحة.');
  console.log('    قد يكون الموقع بحاجة للتحديث أو الحقل مخفي.');
  return null;
}

async function clickSearch(page) {
  const buttons = [
    { name: 'بحث/تدقيق/استعلام/عرض', locator: page.getByRole('button', { name: /بحث|تدقيق|استعلام|عرض/ }).first() },
    { name: 'submit button', locator: page.locator('input[type="submit"], input[type="button"], button').filter({ hasText: /بحث|تدقيق|استعلام|عرض/ }).first() },
  ];

  for (const button of buttons) {
    try {
      const count = await button.locator.count().catch(() => 0);
      if (count > 0 && await button.locator.isVisible()) {
        console.log(`✅ وجدت زر البحث: ${button.name}`);
        await Promise.all([
          page.waitForLoadState('domcontentloaded', { timeout: 5000 }).catch(() => {}),
          button.locator.click(),
        ]);
        await page.waitForTimeout(350);
        return true;
      }
    } catch (err) {
      console.log(`⚠️  محاولة الضغط على ${button.name} فشلت`);
    }
  }

  console.log('⚠️  لم أجد زر بحث في الصفحة.');
  return false;
}


async function searchByNatId(page, natId) {
  const input = await findSearchInput(page);
  if (!input) {
    console.log('❌ لم أتمكن من العثور على حقل البحث.');
    return false;
  }

  try {
    await input.fill(natId);
    console.log(`✏️  أدخلت الرقم: ${natId}`);
  } catch (err) {
    console.log(`❌ فشل إدخال الرقم: ${err.message}`);
    return false;
  }

  return clickSearch(page);
}

async function clearMarks(page) {
  await page.evaluate(() => {
    document.querySelectorAll('[data-karama-current]').forEach((el) => {
      el.style.outline = '';
      el.style.backgroundColor = '';
      el.style.boxShadow = '';
      el.removeAttribute('data-karama-current');
    });
  }).catch(() => {});
}

async function setBanner(page, record, status) {
  await page.evaluate(({ record, status }) => {
    let banner = document.getElementById('karama-update-banner');
    if (!banner) {
      banner = document.createElement('div');
      banner.id = 'karama-update-banner';
      document.body.appendChild(banner);
    }
    banner.textContent = `CURRENT NAT ID: ${record.natId} | ${record.oldAmount} -> ${record.newAmount} | ${status}`;
    banner.style.position = 'fixed';
    banner.style.top = '8px';
    banner.style.left = '8px';
    banner.style.zIndex = '2147483647';
    banner.style.direction = 'ltr';
    banner.style.fontFamily = 'Tahoma, Arial, sans-serif';
    banner.style.fontSize = '18px';
    banner.style.fontWeight = '700';
    banner.style.color = '#111827';
    banner.style.background = '#fde68a';
    banner.style.border = '3px solid #b91c1c';
    banner.style.borderRadius = '6px';
    banner.style.padding = '10px 14px';
    banner.style.boxShadow = '0 10px 30px rgba(0,0,0,.25)';
  }, { record, status }).catch(() => {});
}

async function markRow(page, row, record, status) {
  await clearMarks(page);
  await setBanner(page, record, status);
  await row.evaluate((tr, natId) => {
    tr.setAttribute('data-karama-current', 'row');
    tr.style.outline = '4px solid #dc2626';
    tr.style.backgroundColor = '#fef3c7';
    tr.style.boxShadow = 'inset 0 0 0 9999px rgba(253, 230, 138, 0.35)';
    tr.scrollIntoView({ block: 'center', inline: 'nearest' });

    Array.from(tr.children).forEach((cell) => {
      const digits = String(cell.textContent || '').replace(/[ظ -ظ©]/g, (d) => 'ظ ظ،ظ¢ظ£ظ¤ظ¥ظ¦ظ§ظ¨ظ©'.indexOf(d)).replace(/[^\d]/g, '');
      if (digits === natId) {
        cell.setAttribute('data-karama-current', 'nat-id');
        cell.style.outline = '4px solid #2563eb';
        cell.style.backgroundColor = '#dbeafe';
        cell.style.fontWeight = '700';
      }
    });
  }, record.natId).catch(() => {});
}

async function markAmountInput(input, record, status) {
  await input.evaluate((el, data) => {
    el.setAttribute('data-karama-current', 'amount-input');
    el.style.outline = '4px solid #16a34a';
    el.style.backgroundColor = '#dcfce7';
    el.style.boxShadow = '0 0 0 4px rgba(22, 163, 74, .25)';
    el.title = `CURRENT NAT ID: ${data.natId} | ${data.oldAmount} -> ${data.newAmount} | ${data.status}`;
    el.scrollIntoView({ block: 'center', inline: 'nearest' });
  }, { ...record, status }).catch(() => {});
}

async function findExactNatCells(page, natId) {
  const narrow = page.locator('td, th').filter({ hasText: natId });
  const narrowCount = await narrow.count().catch(() => 0);
  const cells = narrowCount > 0 ? await narrow.all() : await page.locator('td, th').all();
  const exact = [];

  for (const cell of cells) {
    try {
      if (!(await cell.isVisible())) continue;
      const text = await cell.textContent();
      if (normalizeDigits(text) === natId) {
        exact.push(cell);
      }
    } catch (_) {
      // Ignore cells that disappear during postback.
    }
  }

  return exact;
}

async function findVerifiedRow(page, record) {
  for (let pass = 0; pass < 2; pass++) {
    const natCells = await findExactNatCells(page, record.natId);

    for (const natCell of natCells) {
      try {
        const row = natCell.locator('xpath=ancestor::tr[1]');
        if (!(await row.isVisible())) continue;
        const text = await row.textContent();
        if (!text || normalizeDigits(await natCell.textContent()) !== record.natId) continue;
        if (!nameLooksLikeMatch(text, record.name)) {
          console.log(`⚠️  الاسم في الصف لا يطابق الاسم المتوقع:`);
          console.log(`   الاسم المتوقع: "${record.name}"`);
          console.log(`   الاسم في الموقع: "${normalizeText(text).substring(0, 50)}..."`);
          console.log(`   يتم المتابعة للبحث في صفوف أخرى...`);
          continue;
        }

        const inputs = await visibleEditableInputs(row);
        if (inputs.length === 0) {
          console.log(`⚠️  وجدت الصف لكن لم أجد حقول تعديل فيه.`);
          console.log(`   الصف قد يكون للقراءة فقط أو بدون حقول إدخال.`);
          console.log(`   الرقم الوطني: ${record.natId}`);
          continue;
        }

        return { row, natCell };
      } catch (_) {
        // Ignore rows/cells changed during postback.
      }
    }

    if (pass === 0) {
      console.log(`❌ لم أجد الرقم ${record.natId} في الجدول الحالي.`);
      console.log(`🔍 جاري البحث في الموقع...`);
      const searched = await searchByNatId(page, record.natId);
      if (!searched) {
        console.log('⚠️  لم أتمكن من العثور على حقل البحث أو زر البحث.');
      }
    }
  }

  return null;
}

async function findAmountInput(row, oldAmount, newAmount) {
  const inputs = await visibleEditableInputs(row);
  if (inputs.length === 0) {
    return { input: null, reason: 'لا توجد حقول إدخال ظاهرة وقابلة للتعديل داخل الصف' };
  }

  const describedInputs = [];
  for (let i = 0; i < inputs.length; i++) {
    const value = normalizeAmount(await inputs[i].inputValue().catch(() => ''));
    const hints = await describeInputField(inputs[i]);
    describedInputs.push({ i, value, hints, score: amountFieldScore(hints) });
  }

  const oldAmountMatches = oldAmount !== ''
    ? describedInputs.filter((item) => item.value === oldAmount)
    : [];
  if (oldAmountMatches.length === 1) {
    const chosen = oldAmountMatches[0];
    if (chosen.score >= 2) {
      return { input: inputs[chosen.i], reason: `تم اختيار حقل قيمته الحالية ${oldAmount}` };
    }
    return { input: null, reason: `وجدت حقلًا يطابق مبلغ التقرير القديم (${oldAmount}) لكن لا توجد دلائل كافية على أنه حقل المبلغ` };
  }
  if (oldAmountMatches.length > 1) {
    const ranked = [...oldAmountMatches].sort((a, b) => b.score - a.score);
    if (ranked[0].score > ranked[1]?.score) {
      return { input: inputs[ranked[0].i], reason: `تم اختيار أفضل حقل مطابق لمبلغ التقرير القديم ${oldAmount}` };
    }
    return { input: null, reason: `وجدت أكثر من حقل يطابق مبلغ التقرير القديم (${oldAmount}) ولا يوجد تمييز كافٍ بينها` };
  }

  const newAmountMatches = newAmount !== ''
    ? describedInputs.filter((item) => item.value === newAmount)
    : [];
  if (newAmountMatches.length === 1) {
    const chosen = newAmountMatches[0];
    if (chosen.score >= 2) {
      return { input: inputs[chosen.i], alreadyCorrect: true, reason: `المبلغ الحالي في الصفحة يساوي المبلغ المطلوب بالفعل (${newAmount})` };
    }
  }

  if (oldAmount !== '') {
    return { input: null, reason: `لم أجد داخل الصف حقلاً قيمته الحالية تساوي مبلغ التقرير القديم (${oldAmount})` };
  }

  const ranked = [...describedInputs].sort((a, b) => b.score - a.score);
  if (ranked.length > 0 && ranked[0].score >= 6 && ranked[0].score > (ranked[1]?.score ?? -1)) {
    const chosen = ranked[0];
    return {
      input: inputs[chosen.i],
      reason: `تم اختيار حقل يبدو أنه للمبلغ: ${chosen.hints.labelText || chosen.hints.name || chosen.hints.id || chosen.hints.placeholder || 'unknown'}`,
    };
  }

  return { input: null, reason: `لم أستطع تحديد حقل المبلغ بثقة داخل الصف. عدد الحقول: ${inputs.length}` };
}

function isTemporarySaveText(text) {
  const s = normalizeText(text);
  return /حفظ/.test(s) && /مؤقت|مؤقت/.test(s);
}

async function clickTemporarySave(page) {
  const candidates = [
    page.getByRole('button', { name: /حفظ\s*(مؤقت|مؤقت)/i }).first(),
    page.locator('input[type="submit"], input[type="button"], button').filter({ hasText: /حفظ\s*(مؤقت|مؤقت)/i }).first(),
  ];

  for (const candidate of candidates) {
    try {
      if ((await candidate.count()) === 0 || !(await candidate.isVisible())) continue;
      if (!(await candidate.isEnabled().catch(() => false))) continue;

      const label = `${await candidate.getAttribute('value').catch(() => '')} ${await candidate.textContent().catch(() => '')}`;
      if (!isTemporarySaveText(label)) continue;

      await setTemporarySaveBanner(page, 'CLICKING_TEMPORARY_SAVE');
      await Promise.all([
        page.waitForLoadState('domcontentloaded', { timeout: 5000 }).catch(() => {}),
        candidate.click(),
      ]);
      await page.waitForTimeout(500);
      await setTemporarySaveBanner(page, 'TEMPORARY_SAVE_CLICKED');
      return true;
    } catch (_) {
      // Try next candidate.
    }
  }

  const clickedByValue = await page.evaluate(() => {
    const normalize = (value) => String(value || '')
      .trim()
      .replace(/[أإآ]/g, 'ا')
      .replace(/ى/g, 'ي')
      .replace(/ة/g, 'ه')
      .replace(/\s+/g, ' ');
    const isTemporarySave = (value) => /حفظ/.test(normalize(value)) && /مؤقت|مؤقت/.test(normalize(value));
    const controls = Array.from(document.querySelectorAll('input[type="submit"], input[type="button"], button'));
    const el = controls.find((control) => {
      const text = `${control.value || ''} ${control.textContent || ''}`.trim();
      const disabled = control.disabled || control.getAttribute('aria-disabled') === 'true';
      const style = window.getComputedStyle(control);
      return !disabled && style.display !== 'none' && style.visibility !== 'hidden' && isTemporarySave(text);
    });
    if (!el) return false;
    el.scrollIntoView({ block: 'center', inline: 'nearest' });
    el.style.outline = '5px solid #f97316';
    el.style.backgroundColor = '#fed7aa';
    el.click();
    return true;
  }).catch(() => false);

  if (clickedByValue) {
    await page.waitForLoadState('domcontentloaded', { timeout: 5000 }).catch(() => {});
    await page.waitForTimeout(500);
    await setTemporarySaveBanner(page, 'TEMPORARY_SAVE_CLICKED');
    return true;
  }

  return false;
}

async function setTemporarySaveBanner(page, status) {
  await page.evaluate((status) => {
    let banner = document.getElementById('karama-temp-save-banner');
    if (!banner) {
      banner = document.createElement('div');
      banner.id = 'karama-temp-save-banner';
      document.body.appendChild(banner);
    }
    banner.textContent = `TEMPORARY SAVE ONLY: ${status}`;
    banner.style.position = 'fixed';
    banner.style.bottom = '8px';
    banner.style.left = '8px';
    banner.style.zIndex = '2147483647';
    banner.style.direction = 'ltr';
    banner.style.fontFamily = 'Tahoma, Arial, sans-serif';
    banner.style.fontSize = '16px';
    banner.style.fontWeight = '700';
    banner.style.color = '#111827';
    banner.style.background = '#bbf7d0';
    banner.style.border = '3px solid #15803d';
    banner.style.borderRadius = '6px';
    banner.style.padding = '8px 12px';
    banner.style.boxShadow = '0 10px 30px rgba(0,0,0,.25)';
  }, status).catch(() => {});
}

async function openBrowserDirectly(url) {
  return new Promise((resolve) => {
    console.log(`🌐 جاري فتح المتصفح مباشرة: ${url}`);
    const cmd = process.platform === 'win32' 
      ? `start "" "${url}"`
      : process.platform === 'darwin'
      ? `open "${url}"`
      : `xdg-open "${url}"`;
    
    exec(cmd, (error) => {
      if (error) {
        console.log(`⚠️  تحذير: فشل الأمر المباشر، سأحاول Playwright بدلاً منه`);
      } else {
        console.log(`✅ تم فتح المتصفح بنجاح`);
      }
      // دائماً استمر، حتى لو فشل الأمر
      resolve(!error);
    });
    
    // أعطِ الأمر وقتاً لتنفيذه
    setTimeout(resolve, 1000);
  });
}

async function updateValue() {
  console.log(`ملف الإكسل المصدر: ${CONFIG.excelPath}`);
  console.log(`الحفظ المؤقت: ${CONFIG.saveChanges ? `كل ${CONFIG.temporarySaveEvery} تعديل` : 'غير مفعل - سيتم تعديل الحقول فقط'}`);
  console.log(`تعديلات التصفير الى 0: ${CONFIG.allowZeroUpdates ? 'مسموحة' : 'ممنوعة - سيتم تخطيها'}`);
  
  let browser;
  let launchErrors = [];
  const launchOptions = {
    headless: false,
    args: ['--new-window', '--start-maximized'],
  };

  // await openBrowserDirectly(CONFIG.siteUrl);
  // await new Promise(resolve => setTimeout(resolve, 2000));

  try {
    browser = await chromium.launch({ channel: 'chrome', ...launchOptions });
    console.log('Using system Chrome');
  } catch (_chromeErr) {
    launchErrors.push('Chrome: ' + (_chromeErr.message || _chromeErr));
    try {
      browser = await chromium.launch({ channel: 'msedge', ...launchOptions });
      console.log('Using system Edge');
    } catch (_edgeErr) {
      launchErrors.push('Edge: ' + (_edgeErr.message || _edgeErr));
      try {
        browser = await chromium.launch({ ...launchOptions });
        console.log('Using built-in Chromium');
      } catch (_chromiumErr) {
        launchErrors.push('Chromium: ' + (_chromiumErr.message || _chromiumErr));
        throw new Error("تعذر تشغيل اي متصفح.\n\n" + launchErrors.join('\n'));
      }
    }
  }
  const page = await browser.newPage();
  await page.goto(CONFIG.siteUrl, { waitUntil: 'domcontentloaded', timeout: CONFIG.timeoutMs });
  await page.bringToFront();
  await page.waitForTimeout(700);

  console.log('\n╔════════════════════════════════════════╗');
  console.log('║      سجل الدخول للموقع الآن...           ║');
  console.log('║ الواجهة الرسومية سترسل إشارة البدء       ║');
  console.log('╚════════════════════════════════════════╝');
  
  const startCommand = await waitForGUICommandObject();
  if (startCommand.action !== 'start') {
    console.log('لم يتم استقبال إشارة البدء. جاري الخروج.');
    return;
  }
  console.log('✅ تم استقبال إشارة البدء - في انتظار أول سجل للمعالجة...');

  const results = [];
  let changedSinceSave = 0;
  let totalChanged = 0;

  mainLoop:
  while(true) {
    console.log('\n' + '='.repeat(70));
    console.log('⏳ في انتظار السجل التالي من الواجهة الرسومية...');
    const command = await waitForGUICommandObject();

    let record = command.record;

    switch (command.action) {
      case 'shutdown':
        console.log('💤 أمر بالإغلاق مستلم. جاري إنهاء العملية...');
        break mainLoop;

      case 'skip':
        console.log(`⏭️  تم استلام أمر التخطي للسجل: ${record?.natId}`);
        if(record) {
          results.push({ ...record, status: 'skipped_by_user' });
        }
        continue mainLoop;

      case 'process':
        if (!record) {
          console.log('⚠️  أمر المعالجة لا يحتوي على بيانات السجل.');
          continue mainLoop;
        }

        console.log(`[${command.index + 1}/${command.total}] NAT=${record.natId} | AMOUNT ${record.oldAmount} -> ${record.newAmount}`);
        writeCurrentStatus(record, 'START_SEARCH');
        await setBanner(page, record, 'START_SEARCH');

        try {
          const targetAmount = normalizeAmount(record.newAmount);
          if (targetAmount === '0' && !CONFIG.allowZeroUpdates) {
            throw new Error(`التعديل إلى 0 ممنوع. استخدم --allow-zero للسماح بذلك.`);
          }

          console.log(`STEP 1: 🔍 البحث عن الرقم الوطني: ${record.natId}`);
          const rowData = await findVerifiedRow(page, record);
          if (!rowData || !rowData.row) {
            throw new Error(`❌ لم أجد الشخص ذا الرقم الوطني ${record.natId} مع الاسم "${record.name}"`);
          }
          const { row } = rowData;

          const rowText = normalizeText(await row.textContent());
          const natCellDigits = normalizeDigits(await rowData.natCell.textContent().catch(() => ''));
          if (natCellDigits !== record.natId || !nameLooksLikeMatch(rowText, record.name)) {
            throw new Error('فشل التحقق النهائي من الرقم الوطني والاسم قبل التعديل');
          }
          console.log(`STEP 2: Verified NAT and name for NAT=${record.natId}`);
          writeCurrentStatus(record, 'ROW_VERIFIED');
          await markRow(page, row, record, 'ROW_VERIFIED');

          const { input, reason, alreadyCorrect } = await findAmountInput(row, record.oldAmount, targetAmount);
          if (!input) {
            throw new Error(reason);
          }

          const before = normalizeAmount(await input.inputValue().catch(() => ''));
          console.log(`STEP 3: Amount input selected. Current=${before || '(empty)'}`);
          console.log(`FIELD_REASON: ${reason}`);
          writeCurrentStatus({ ...record, newAmount: targetAmount }, `AMOUNT_FIELD_SELECTED current=${before || '(empty)'}`);
          await markAmountInput(input, record, 'AMOUNT_FIELD_SELECTED');

          if (alreadyCorrect) {
            const status = 'ALREADY_CORRECT';
            results.push({ ...record, status, current: before });
            writeCurrentStatus(record, status);
            await setBanner(page, record, status);
            console.log(`✅ القيمة صحيحة بالفعل: ${before} = ${record.newAmount}`);
            continue mainLoop;
          }

          if (record.oldAmount && before && before !== record.oldAmount) {
            throw new Error(`قيمة الحقل الحالية (${before}) لا تطابق قيمة التقرير (${record.oldAmount})`);
          }

          if (CONFIG.requireConfirmBeforeEdit || CONFIG.confirmEach) {
            await markRow(page, row, record, 'WAITING_USER_CONFIRMATION');
            await markAmountInput(input, record, 'WAITING_USER_CONFIRMATION');
            writeCurrentStatus(record, 'WAITING_USER_CONFIRMATION');
            
            console.log('⏳ في انتظار تأكيد المستخدم من الواجهة الرسومية...');
            const ans = await ask(`VERIFY highlighted row. Edit NAT=${record.natId} AMOUNT ${before} -> ${targetAmount}?`);
            if (ans !== 'yes') {
              throw new Error(`تم التخطي بواسطة المستخدم للرقم الوطني ${record.natId}`);
            }
          }

          await input.fill('');
          await input.fill(targetAmount);
          console.log(`✏️  تم التعديل: ${before || '(فارغ)'} → ${targetAmount}`);
          writeCurrentStatus(record, 'FIELD_CHANGED');
          await markRow(page, row, record, 'FIELD_CHANGED');
          await markAmountInput(input, record, 'FIELD_CHANGED');

          totalChanged += 1;
          changedSinceSave += 1;

          let saved = false;
          if (CONFIG.saveChanges && changedSinceSave >= CONFIG.temporarySaveEvery) {
            console.log(`💾 حفظ مؤقت: تم ${changedSinceSave} تعديل. جاري الحفظ...`);
            writeCurrentStatus(record, `TEMP_SAVE_AFTER_${changedSinceSave}_CHANGES`);
            saved = await clickTemporarySave(page);
            if (saved) {
              changedSinceSave = 0;
              console.log('✅ تم الحفظ المؤقت بنجاح.');
            } else {
              console.log('⚠️  لم أجد زر الحفظ المؤقت أو لم يكن مفعلاً.');
            }
          } else if (!CONFIG.saveChanges) {
            console.log('ℹ️  الحفظ معطّل - التعديل على الشاشة فقط.');
          } else {
            console.log(`⏳ في الانتظار: ${changedSinceSave}/${CONFIG.temporarySaveEvery} تعديل قبل الحفظ.`);
          }

          const status = saved ? 'FIELD_CHANGED_TEMP_SAVED' : 'FIELD_CHANGED_PENDING_SAVE';
          results.push({ ...record, status, current: before, newAmount: targetAmount });
          writeCurrentStatus(record, status);
          await setBanner(page, record, status);
          console.log(saved ? '✅ تم التعديل والحفظ.' : '⏳ تم التعديل - في الانتظار للحفظ.');

        } catch (err) {
          results.push({ ...record, status: 'error', error: err.message });
          writeCurrentStatus(record, `ERROR: ${err.message}`);
          await setBanner(page, record, `ERROR: ${err.message}`);
          console.log(`❌ خطأ: ${err.message}`);
        }
        break; 

      default:
        console.log(`⚠️  أمر غير معروف: ${command.action}`);
        break;
    }
  }

  if (CONFIG.saveChanges && changedSinceSave > 0) {
    console.log(`\n💾 حفظ مؤقت نهائي: جاري حفظ ${changedSinceSave} تعديل متبقية...`);
    const finalSaved = await clickTemporarySave(page);
    console.log(finalSaved ? '✅ تم الحفظ النهائي بنجاح.' : '⚠️  لم يتمكن من إكمال الحفظ النهائي.');
  }

  const outPath = outputPath(`auto_update_result_${Date.now()}.json`);
  fs.writeFileSync(outPath, JSON.stringify(results, null, 2), 'utf8');
  console.log(`\n✅ انتهت العملية بنجاح!`);
  console.log(`📊 التقرير: ${outPath}`);
  console.log(`📝 عدد التعديلات: ${totalChanged}`);

  if (CONFIG.fromGUI) {
    console.log('__DONE__'); // Signal to GUI that we are done
    console.log(`🔔 تم الانتهاء من التعديلات! اضغط حفظ مؤقت في الموقع لحفظ التغييرات.`);
    console.log(`ℹ️  اترك نافذة المتصفح مفتوحة للمراجعة.`);
    await new Promise(() => {}); // Hang forever to keep browser open
  } else {
    console.log(`ℹ️  جاري إغلاق المتصفح...`);
    await browser.close();
    console.log(`🚪 تم إغلاق المتصفح.`);
  }
}

updateValue().then(async () => {
  console.log('\n✅ اكتمل السكريبت بنجاح.');
  // Do not exit when successful, let the updateValue function handle it,
  // especially for GUI mode where we want the browser to stay open.
}).catch(async (err) => {
  console.error('\n========================================');
  console.error('ERROR: ' + (err.message || err));
  console.error('========================================\n');
  console.error('النافذة ستغلق تلقائيا بعد دقيقتين...');
  await new Promise(resolve => setTimeout(resolve, 120000));
  process.exit(1);
});
