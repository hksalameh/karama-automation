const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const xlsx = require('xlsx');

const LOGIN_URL = 'http://92.253.101.105:8040/ICCS/index.aspx';
const DISB_URL = '/ICCS/Family_sub2.aspx';
const FAIL_TIMEOUT = 25000; // حد فشل فقط، وليس انتظاراً ثابتاً
const POLL = 80;

const sleep = (ms) => new Promise(r => setTimeout(r, ms));
const normText = v => String(v ?? '').trim().replace(/[أإآ]/g,'ا').replace(/ى/g,'ي').replace(/ة/g,'ه').replace(/ـ/g,'').replace(/\s+/g,' ');
const normDigits = v => String(v ?? '').trim().replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/\D/g,'');
const normAmount = v => { const s=String(v??'').trim().replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/,/g,'').replace(/[^\d.-]/g,''); if(!s)return ''; const n=Number(s); return Number.isFinite(n)?String(n):s; };
const log = m => console.log(m);
function progress(i,t,r,s){ const safe=v=>String(v??'').replace(/[|\r\n]/g,' '); console.log(`__PROGRESS__|${i}|${t}|${safe(r.natId)}|${safe(r.oldAmount)}|${safe(r.newAmount)}|${safe(s)}`); }
function outPath(name){ const d=process.env.KARAMA_OUTPUT_DIR||process.cwd(); fs.mkdirSync(d,{recursive:true}); return path.join(d,name); }

async function waitUntil(fn, desc, timeout=FAIL_TIMEOUT){
  const end=Date.now()+timeout; let last='';
  while(Date.now()<end){ try{ const v=await fn(); if(v)return v; }catch(e){ last=e?.message||String(e); } await sleep(POLL); }
  throw new Error(`انتهت مهلة انتظار ${desc}${last?`: ${last}`:''}.`);
}

function config(){
  const c={excelPath:process.env.KARAMA_EXCEL_PATH||'',username:process.env.KARAMA_USERNAME||'',password:process.env.KARAMA_PASSWORD||'',year:String(process.env.KARAMA_YEAR||'').trim(),month:String(process.env.KARAMA_MONTH||'').trim(),category:String(process.env.KARAMA_CATEGORY||'').trim()};
  if(!c.excelPath||!fs.existsSync(c.excelPath)) throw new Error(`ملف نتيجة المقارنة غير موجود: ${c.excelPath||'(فارغ)'}`);
  if(!c.username||!c.password) throw new Error('اسم المستخدم أو كلمة المرور غير مدخلين.');
  if(!/^\d{4}$/.test(c.year)) throw new Error(`السنة غير صحيحة: ${c.year}`);
  const m=Number(c.month); if(!Number.isInteger(m)||m<1||m>12) throw new Error(`الشهر غير صحيح: ${c.month}`); c.month=String(m);
  if(!c.category) throw new Error('التصنيف غير محدد.'); return c;
}

function loadRecords(file){
  const wb=xlsx.readFile(file,{cellDates:false});
  const sheet=wb.SheetNames.find(n=>normText(n).includes('يحتاج تعديل'))||wb.SheetNames.find(n=>normText(n).includes('تعديل'));
  if(!sheet) throw new Error(`لم يتم العثور على ورقة "يحتاج تعديل". الأوراق: ${wb.SheetNames.join(', ')}`);
  const rows=xlsx.utils.sheet_to_json(wb.Sheets[sheet],{defval:'',raw:false}); if(!rows.length)return [];
  const headers=Object.keys(rows[0]);
  const pick=pred=>headers.find(h=>pred(normText(h)));
  const nat=pick(h=>h.includes('وطني')), name=pick(h=>h.includes('اسم')&&h.includes('الموقع')), old=pick(h=>h.includes('مبلغ')&&h.includes('الموقع')), neu=pick(h=>h.includes('مبلغ')&&(h.includes('فعلي')||h.includes('الجديد')||h.includes('متوقع'))), reason=pick(h=>h.includes('سبب'));
  if(!nat||!old||!neu) throw new Error(`تعذر اكتشاف أعمدة ملف النتيجة. الأعمدة: ${headers.join(' | ')}`);
  return rows.map(r=>({natId:normDigits(r[nat]),name:String(name?r[name]:'').trim(),oldAmount:normAmount(r[old]),newAmount:normAmount(r[neu]),reason:String(reason?r[reason]:'').trim()})).filter(r=>r.natId&&r.newAmount!=='');
}

async function launch(){
  const tries=[['Chrome',{channel:'chrome',headless:false}],['Edge',{channel:'msedge',headless:false}],['Chromium',{headless:false}]], errors=[];
  for(const [n,o] of tries){ try{ const b=await chromium.launch(o); log(`✅ تم تشغيل المتصفح: ${n}`); return b; }catch(e){ errors.push(`${n}: ${e.message||e}`); } }
  throw new Error(`تعذر تشغيل Chrome أو Edge.\n${errors.join('\n')}`);
}

async function label(el){ return el.evaluate(e=>`${e.value||''} ${e.innerText||e.textContent||''}`.trim()).catch(()=> ''); }
const finalSave = t => { const n=normText(t); return n.includes('حفظ')&&n.includes('نهائي'); };
async function clickable(page,text,exact=true){
  const wanted=normText(text), els=page.locator('button,input[type="submit"],input[type="button"],a');
  for(let i=0;i<await els.count();i++){ const e=els.nth(i); try{ if(!await e.isVisible()||!await e.isEnabled())continue; const l=normText(await label(e)); if(finalSave(l))continue; if((exact&&l===wanted)||(!exact&&l.includes(wanted)))return e; }catch(_){} }
  return null;
}
async function bodyHas(page,arr){ const b=normText(await page.locator('body').innerText().catch(()=>'')); return (Array.isArray(arr)?arr:[arr]).some(x=>b.includes(normText(x))); }
async function cycle(page){ return page.evaluate(()=>{ const v=n=>document.querySelector(`input[name="${n}"]`)?.value||''; return `${performance.timeOrigin}|${location.href}|${v('__VIEWSTATE').slice(-160)}|${v('__EVENTVALIDATION').slice(-160)}`; }).catch(()=> ''); }
async function waitCycle(page,before,desc){ return waitUntil(async()=>{ const now=await cycle(page); return now&&now!==before; },desc); }

async function login(page,c){
  log('🌐 فتح صفحة الدخول...'); await page.goto(LOGIN_URL,{waitUntil:'domcontentloaded',timeout:FAIL_TIMEOUT});
  const pass=page.locator('input[type="password"]:visible').first(); await waitUntil(()=>pass.isVisible().catch(()=>false),'ظهور حقل كلمة المرور');
  const texts=page.locator('input[type="text"]:visible,input:not([type]):visible'); let user=null;
  for(let i=0;i<await texts.count();i++){ if(await texts.nth(i).isVisible().catch(()=>false)){ user=texts.nth(i); break; } }
  if(!user)throw new Error('لم أجد حقل اسم المستخدم.'); await user.fill(c.username); await pass.fill(c.password);
  const btn=await clickable(page,'دخول'); if(!btn)throw new Error('لم أجد زر "دخول".'); log('⏳ تسجيل الدخول...'); await btn.click({timeout:FAIL_TIMEOUT});
  await waitUntil(async()=>await bodyHas(page,['القائمة الرئيسية','شاشة المعيل','شاشة المستفيدين'])||!page.url().toLowerCase().endsWith('/index.aspx'),'ظهور الصفحة الرئيسية'); log('✅ تم تسجيل الدخول.');
}

async function goDisbursement(page){
  if(page.url().includes(DISB_URL))return; log('➡️ الدخول إلى شاشة الصرفية...'); const btn=await clickable(page,'شاشة الصرفية'); if(!btn)throw new Error('لم أجد خيار "شاشة الصرفية".');
  await btn.click({timeout:FAIL_TIMEOUT}); await waitUntil(async()=>page.url().includes('Family_sub2.aspx')||await bodyHas(page,['شاشة الصرفيات','شاشه الصرفيات']),'ظهور شاشة الصرفية'); log('✅ تم فتح شاشة الصرفية.');
}

async function selectInfo(sel){ return sel.evaluate(e=>({options:Array.from(e.options||[]).map(o=>(o.textContent||'').trim()),parent:(e.closest('tr')?.innerText||e.parentElement?.innerText||'').trim()})).catch(()=>({options:[],parent:''})); }
async function monthSelect(page){ const ss=page.locator('select:visible'); let best=null,score=-1; for(let i=0;i<await ss.count();i++){ const s=ss.nth(i),d=await selectInfo(s),vals=new Set(d.options.map(normDigits).filter(Boolean)); let x=[...Array(12)].every((_,k)=>vals.has(String(k+1)))?8:0; if(normText(d.parent).includes('شهر'))x+=4; if(x>score){score=x;best=s;} } return score>=8?best:null; }
async function categorySelect(page){ const ss=page.locator('select:visible'); let best=null,score=-1; for(let i=0;i<await ss.count();i++){ const s=ss.nth(i),d=await selectInfo(s),o=d.options.map(normText); let x=0; if(o.some(v=>v.includes('ايتام')))x+=4; if(o.some(v=>v.includes('اسر')))x+=4; if(o.some(v=>v.includes('طلاب علم')))x+=4; if(normText(d.parent).includes('تصنيف الاسره'))x+=5; if(x>score){score=x;best=s;} } return score>=4?best:null; }
async function yearInput(page){ const ins=page.locator('input:visible'); let best=null,score=-1; for(let i=0;i<await ins.count();i++){ const e=ins.nth(i); try{ const type=((await e.getAttribute('type'))||'text').toLowerCase(); if(/hidden|submit|button|password|checkbox|radio|file/.test(type))continue; const val=normDigits(await e.inputValue().catch(()=>'')),ctx=normText(await e.evaluate(x=>x.closest('tr')?.innerText||x.parentElement?.innerText||'').catch(()=>'')); let s=/^\d{4}$/.test(val)?5:0; if(ctx.includes('سنه')||ctx.includes('سنة'))s+=6; if(s>score){score=s;best=e;} }catch(_){} } return score>=5?best:null; }
async function chooseMonth(sel,wanted){ for(const o of await sel.locator('option').all()){ if(normDigits(await o.textContent().catch(()=>''))===wanted){ const v=await o.getAttribute('value'); if(v!==null){ await sel.selectOption(v); return true; } } } return false; }
async function chooseCategory(sel,wanted){ const w=normText(wanted); for(const o of await sel.locator('option').all()){ const t=normText(await o.textContent().catch(()=>'')); if(t===w||t.includes(w)||w.includes(t)){ const v=await o.getAttribute('value'); if(v!==null){ await sel.selectOption(v); return true; } } } return false; }

async function screenState(page){
  return page.evaluate(()=>{
    const nt=v=>String(v||'').trim().replace(/[أإآ]/g,'ا').replace(/ى/g,'ي').replace(/ة/g,'ه').replace(/ـ/g,'').replace(/\s+/g,' '), dg=v=>String(v||'').replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/\D/g,'');
    const vis=e=>{const s=getComputedStyle(e),r=e.getBoundingClientRect();return s.display!=='none'&&s.visibility!=='hidden'&&r.width>0&&r.height>0;}; let month='',category='',year='';
    for(const s of document.querySelectorAll('select')){ if(!vis(s))continue; const opts=Array.from(s.options||[]).map(o=>nt(o.textContent||'')), vals=new Set(opts.map(dg).filter(Boolean)),selected=nt(s.options?.[s.selectedIndex]?.textContent||''); if(!month&&Array.from({length:12},(_,i)=>String(i+1)).every(m=>vals.has(m)))month=dg(selected); if(!category&&opts.some(x=>x.includes('ايتام'))&&opts.some(x=>x.includes('اسر')))category=selected; }
    for(const i of document.querySelectorAll('input')){ if(!vis(i)||/hidden|submit|button|password|checkbox|radio|file/.test(String(i.type||'text')))continue; const v=dg(i.value),ctx=nt(i.closest('tr')?.innerText||i.parentElement?.innerText||''); if(/^\d{4}$/.test(v)&&(ctx.includes('سنه')||ctx.includes('سنة'))){year=v;break;} }
    let rows=0; for(const tr of document.querySelectorAll('tr')){ const cells=Array.from(tr.cells||[]); if(cells.length<2)continue; const hasNat=cells.some(c=>/^\d{8,12}$/.test(dg(c.innerText||c.textContent||''))); if(!hasNat)continue; const ins=Array.from(tr.querySelectorAll(':scope > td input,:scope > th input')).filter(i=>!i.disabled&&!i.readOnly&&!/hidden|button|submit|checkbox|radio|image/.test(String(i.type||'text'))&&vis(i)); if(ins.length)rows++; }
    return {month,category,year,rows};
  }).catch(()=>({month:'',category:'',year:'',rows:0}));
}

async function configure(page,c){
  log(`⚙️ تجهيز الصرفية: سنة ${c.year}، شهر ${c.month}، تصنيف ${c.category}`); const y=await yearInput(page); if(!y)throw new Error('لم أجد حقل السنة.'); await y.fill(c.year);
  const m=await monthSelect(page); if(!m||!await chooseMonth(m,c.month))throw new Error(`لم أجد الشهر ${c.month}.`); const cat=await categorySelect(page); if(!cat||!await chooseCategory(cat,c.category))throw new Error(`لم أجد التصنيف "${c.category}".`);
  const show=await clickable(page,'عرض'); if(!show)throw new Error('لم أجد زر "عرض".'); const before=await cycle(page); log('⏳ عرض الصرفية...'); await show.click({timeout:FAIL_TIMEOUT}); await waitCycle(page,before,'تحديث صفحة الصرفية بعد عرض');
  const wanted=normText(c.category); const state=await waitUntil(async()=>{ const s=await screenState(page),sc=normText(s.category); if(s.rows>0&&s.month===c.month&&(!s.year||s.year===c.year)&&(!sc||sc===wanted||sc.includes(wanted)||wanted.includes(sc)))return s; return null; },'ظهور صفوف الصرفية');
  log(`✅ ظهرت الصرفية وأصبحت جاهزة (${state.rows} صف قابل للتعديل).`);
}

async function installHelpers(page){
  await page.evaluate(()=>{
    window.__karama={}; const K=window.__karama;
    K.nt=v=>String(v||'').trim().replace(/[أإآ]/g,'ا').replace(/ى/g,'ي').replace(/ة/g,'ه').replace(/ـ/g,'').replace(/\s+/g,' ');
    K.dg=v=>String(v||'').replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/\D/g,'');
    K.am=v=>{const s=String(v??'').trim().replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/,/g,'').replace(/[^\d.-]/g,'');if(!s)return '';const n=Number(s);return Number.isFinite(n)?String(n):s;};
    K.vis=e=>{const s=getComputedStyle(e),r=e.getBoundingClientRect();return s.display!=='none'&&s.visibility!=='hidden'&&r.width>0&&r.height>0;};
    K.rows=id=>Array.from(document.querySelectorAll('tr')).filter(tr=>Array.from(tr.cells||[]).some(c=>K.dg(c.innerText||c.textContent||'')===id));
    K.inputs=tr=>Array.from(tr.querySelectorAll(':scope > td input,:scope > th input')).filter(i=>!i.disabled&&!i.readOnly&&!/hidden|button|submit|checkbox|radio|image/.test(String(i.type||'text'))&&K.vis(i));
    K.valueInput=(tr,oldVal,newVal)=>{ const ins=K.inputs(tr),old=ins.filter(i=>K.am(i.value)===oldVal); if(old.length===1)return old[0]; const neu=ins.filter(i=>K.am(i.value)===newVal); if(neu.length===1)return neu[0]; const table=tr.closest('table'); if(table){ for(const hr of Array.from(table.rows||[])){ const cells=Array.from(hr.cells||[]),txt=cells.map(c=>K.nt(c.innerText||c.textContent||'')); if(!txt.some(t=>t.includes('الرقم الوطني')))continue; const ix=txt.findIndex(t=>t==='القيمه'||t==='القيمة'||t.includes('القيمه')||t.includes('القيمة')); if(ix>=0&&ix<tr.cells.length){const ci=Array.from(tr.cells[ix].querySelectorAll('input')).filter(i=>!i.disabled&&!i.readOnly&&K.vis(i));if(ci.length===1)return ci[0];} } } const nums=ins.filter(i=>K.am(i.value)!==''); return nums.length===1?nums[0]:null; };
    K.check=r=>{ const rows=K.rows(r.natId); if(rows.length!==1)return rows.length?`الرقم ${r.natId} موجود في أكثر من صف`:`الرقم ${r.natId} غير موجود في الصرفية`; const tr=rows[0],expected=K.nt(r.name); if(expected){const text=K.nt(tr.innerText||tr.textContent||''),tokens=expected.split(' ').filter(x=>x.length>=3),matched=tokens.filter(x=>text.includes(x)).length;if(!text.includes(expected)&&tokens.length&&matched<Math.min(3,tokens.length))return `الاسم لا يطابق الرقم ${r.natId}`;} const old=K.am(r.oldAmount),neu=K.am(r.newAmount),input=K.valueInput(tr,old,neu); if(!input)return `تعذر تحديد خانة القيمة للرقم ${r.natId}`; const cur=K.am(input.value); if(cur!==old&&cur!==neu)return `القيمة الحالية ${cur} للرقم ${r.natId} لا تطابق القديمة ${old} ولا الجديدة ${neu}`; return ''; };
    K.apply=r=>{ const err=K.check(r); if(err)return {ok:false,error:err}; const tr=K.rows(r.natId)[0],old=K.am(r.oldAmount),neu=K.am(r.newAmount),input=K.valueInput(tr,old,neu),cur=K.am(input.value); if(cur===neu)return {ok:true,status:'already_correct',current:cur}; const d=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value'); if(d?.set)d.set.call(input,neu);else input.value=neu; input.dispatchEvent(new Event('input',{bubbles:true})); input.dispatchEvent(new Event('change',{bubbles:true})); const after=K.am(input.value); return after===neu?{ok:true,status:'changed',current:cur,after}:{ok:false,error:`لم تثبت القيمة الجديدة للرقم ${r.natId}`}; };
  });
}
async function preflight(page,records){ return page.evaluate(rs=>rs.map((r,i)=>({i,natId:r.natId,error:window.__karama.check(r)})).filter(x=>x.error),records); }
async function processRecord(page,r,i,total){ progress(i,total,r,'جاري التعديل'); const x=await page.evaluate(r=>window.__karama.apply(r),r); if(!x.ok)throw new Error(x.error); if(x.status==='already_correct'){progress(i,total,r,'صحيح مسبقاً');return {...r,status:'already_correct',current:x.current};} progress(i,total,r,'تم التعديل');return {...r,status:'changed',current:x.current,newAmount:x.after}; }

async function tempSave(page){
  log('💾 انتهت جميع التعديلات. جاري تنفيذ حفظ مؤقت فقط...'); const btn=await clickable(page,'حفظ مؤقت'); if(!btn)throw new Error('لم أجد زر "حفظ مؤقت". لم يتم تنفيذ أي حفظ نهائي.'); if(finalSave(await label(btn)))throw new Error('حماية الأمان منعت زر حفظ نهائي.');
  const before=await cycle(page); await btn.click({timeout:FAIL_TIMEOUT}); await waitCycle(page,before,'استجابة الموقع بعد حفظ مؤقت'); await waitUntil(()=>bodyHas(page,'تم حفظ القيم بنجاح'),'ظهور رسالة نجاح الحفظ المؤقت'); log('✅ تم الحفظ المؤقت وظهرت رسالة النجاح.');
}
async function visibleTotal(page){ return page.evaluate(()=>{const n=v=>String(v||'').trim().replace(/\s+/g,' '),nodes=Array.from(document.querySelectorAll('td,th,span,label,div')),l=nodes.find(e=>n(e.innerText||e.textContent||'').includes('المجموع الكلي للصرفية'));if(!l)return '';const r=l.closest('tr')||l.parentElement,i=r?.querySelector('input');if(i?.value)return i.value;const m=n(r?.innerText||r?.textContent||'').match(/المجموع الكلي للصرفية\s*([\d.,]+)/);return m?m[1]:'';}).catch(()=>''); }

async function main(){
  const c=config(),records=loadRecords(c.excelPath); if(!records.length)throw new Error('لا توجد سجلات في ورقة "يحتاج تعديل".'); log(`📄 تم تحميل ${records.length} سجل يحتاج تعديل.`);
  const dup=records.map(r=>r.natId).filter((id,i,a)=>a.indexOf(id)!==i); if(dup.length)throw new Error(`يوجد رقم وطني مكرر في ملف التعديل: ${[...new Set(dup)].join(', ')}`);
  const browser=await launch(),page=await browser.newPage(); page.setDefaultTimeout(FAIL_TIMEOUT); page.setDefaultNavigationTimeout(FAIL_TIMEOUT); const results=[]; let temporarySaved=false;
  try{
    await login(page,c); await goDisbursement(page); await configure(page,c); await installHelpers(page);
    log('🔎 فحص جميع السجلات قبل بدء أي تعديل...'); const errors=await preflight(page,records); if(errors.length){const p=errors.slice(0,8).map(e=>`${e.natId}: ${e.error}`).join(' | ');throw new Error(`فشل فحص الأمان قبل التعديل (${errors.length} سجل): ${p}`);} log('✅ تم التحقق من جميع السجلات. بدء التعديل السريع داخل الصفحة.');
    for(let i=0;i<records.length;i++){ try{results.push(await processRecord(page,records[i],i+1,records.length));}catch(e){results.push({...records[i],status:'error',error:e.message||String(e)});fs.writeFileSync(outPath(`auto_update_result_${Date.now()}.json`),JSON.stringify(results,null,2),'utf8');throw new Error(`توقفت العملية عند السجل ${i+1}/${records.length}: ${e.message||e}\nلم يتم الضغط على حفظ مؤقت.`);} }
    await tempSave(page); temporarySaved=true; const total=await visibleTotal(page),report=outPath(`auto_update_result_${Date.now()}.json`); fs.writeFileSync(report,JSON.stringify({year:c.year,month:c.month,category:c.category,temporarySaved:true,visibleTotal:total,results},null,2),'utf8'); console.log(`__TEMP_SAVE_OK__|${total||''}|${report}`); log('🛑 انتهت مسؤولية البرنامج هنا. لن يتم الضغط على "حفظ نهائي" بأي شكل.'); log('👀 المتصفح سيبقى مفتوحاً للمراجعة والحفظ النهائي اليدوي.'); while(browser.isConnected())await sleep(1000);
  }catch(e){fs.writeFileSync(outPath(`auto_update_error_${Date.now()}.json`),JSON.stringify({year:c.year,month:c.month,category:c.category,temporarySaved,error:e.message||String(e),results},null,2),'utf8');throw e;}
}
main().catch(e=>{const m=String(e.message||e).replace(/[\r\n]+/g,' ');console.error(`__FATAL__|${m}`);console.error(`❌ ${m}`);process.exitCode=1;});
