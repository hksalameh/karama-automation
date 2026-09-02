# -*- coding: utf-8 -*-
"""Apply idempotent browser compatibility fixes to Karama automation before build."""
from pathlib import Path

PATH = Path(__file__).with_name('auto_update_from_diff.js')
text = PATH.read_text(encoding='utf-8')
changed = False

# ---------------------------------------------------------------------------
# 1) Robust year-field detection.
# ---------------------------------------------------------------------------
YEAR_MARKER = '// YEAR_FIELD_ROBUST_V3'
if YEAR_MARKER not in text:
    old_prepare = """    for(const i of document.querySelectorAll('input')){\n      if(!vis(i)||/hidden|submit|button|password|checkbox|radio|file/.test(String(i.type||'text').toLowerCase()))continue;\n      const val=dg(i.value),ctx=nt(i.closest('tr')?.innerText||i.parentElement?.innerText||'');\n      if((/^\\d{4}$/.test(val)||ctx.includes('سنه')||ctx.includes('سنة'))&&(ctx.includes('سنه')||ctx.includes('سنة'))){yearInp=i;break;}\n    }\n"""

    new_prepare = """    // YEAR_FIELD_ROBUST_V3\n    // حقل السنة في نسخ موقع كرامة قد يكون input أو select، وقد لا تكون التسمية داخل نفس <tr>.\n    const usableInputs=Array.from(document.querySelectorAll('input')).filter(i=>\n      vis(i)&&!/hidden|submit|button|password|checkbox|radio|file/.test(String(i.type||'text').toLowerCase())\n    );\n    const usableSelects=Array.from(document.querySelectorAll('select')).filter(s=>vis(s));\n    const controlDigits=e=>{\n      if(e.tagName==='SELECT'){\n        const opt=e.options?.[e.selectedIndex];\n        return dg((opt?.textContent||'')+' '+(e.value||''));\n      }\n      return dg(e.value||'');\n    };\n    const plausibleYear=e=>{ const v=controlDigits(e); return /^\\d{4}$/.test(v)&&Number(v)>=1900&&Number(v)<=2200; };\n    const allYearCandidates=[...usableInputs,...usableSelects].filter(plausibleYear);\n    if(allYearCandidates.length){\n      yearInp=allYearCandidates.find(e=>controlDigits(e)===String(year))||null;\n      if(!yearInp&&allYearCandidates.length===1)yearInp=allYearCandidates[0];\n      if(!yearInp){\n        const yearLabels=Array.from(document.querySelectorAll('label,td,th,span,div')).filter(e=>{\n          if(!vis(e))return false;\n          const t=nt(e.innerText||e.textContent||'');\n          return t==='السنه'||t==='السنة'||t.includes('السنه')||t.includes('السنة');\n        });\n        let best=null,bestDistance=Infinity;\n        for(const e of allYearCandidates){\n          const er=e.getBoundingClientRect();\n          for(const l of yearLabels){\n            const lr=l.getBoundingClientRect();\n            const d=Math.abs((er.top+er.height/2)-(lr.top+lr.height/2))*5+Math.abs((er.left+er.width/2)-(lr.left+lr.width/2));\n            if(d<bestDistance){bestDistance=d;best=e;}\n          }\n        }\n        yearInp=best;\n      }\n    }\n    if(!yearInp){\n      yearInp=[...usableInputs,...usableSelects].find(e=>/year|yr|sana|sanah|سنه|سنة/i.test(`${e.id||''} ${e.name||''} ${e.getAttribute('aria-label')||''}`))||null;\n    }\n"""

    old_state = """    for(const i of document.querySelectorAll('input')){ if(!vis(i)||/hidden|submit|button|password|checkbox|radio|file/.test(String(i.type||'text')))continue; const v=dg(i.value),ctx=nt(i.closest('tr')?.innerText||i.parentElement?.innerText||''); if(/^\\d{4}$/.test(v)&&(ctx.includes('سنه')||ctx.includes('سنة'))){year=v;break;} }\n"""
    new_state = """    for(const e of document.querySelectorAll('input,select')){ if(!vis(e))continue; if(e.tagName==='INPUT'&&/hidden|submit|button|password|checkbox|radio|file/.test(String(e.type||'text')))continue; const raw=e.tagName==='SELECT'?`${e.options?.[e.selectedIndex]?.textContent||''} ${e.value||''}`:(e.value||''); const v=dg(raw); if(/^\\d{4}$/.test(v)&&Number(v)>=1900&&Number(v)<=2200){year=v;break;} }\n"""

    if old_prepare not in text:
        raise SystemExit('Expected prepareFiltersDirect year-detection block was not found; refusing silent patch.')
    if old_state not in text:
        raise SystemExit('Expected screenState year-detection block was not found; refusing silent patch.')

    text = text.replace(old_prepare, new_prepare, 1)
    text = text.replace(old_state, new_state, 1)

    old_error = "    if(!yearInp)return {ok:false,error:'لم أجد حقل السنة.'};\n"
    new_error = """    if(!yearInp){\n      const diag=[...usableInputs,...usableSelects].slice(0,20).map(e=>{\n        const kind=e.tagName.toLowerCase();\n        const id=e.id||''; const name=e.name||'';\n        const digits=kind==='select'?dg(`${e.options?.[e.selectedIndex]?.textContent||''} ${e.value||''}`):dg(e.value||'');\n        return `${kind}#${id}[${name}]${digits?`=${digits}`:''}`;\n      }).join(' | ');\n      return {ok:false,error:`لم أجد حقل السنة. الحقول المرئية: ${diag}`};\n    }\n"""
    if old_error not in text:
        raise SystemExit('Expected year failure message was not found; refusing silent patch.')
    text = text.replace(old_error, new_error, 1)

    old_set = "    yearInp.value=String(year);\n"
    new_set = """    if(yearInp.tagName==='SELECT'){\n      const yOpt=Array.from(yearInp.options||[]).find(o=>dg(`${o.textContent||''} ${o.value||''}`)===String(year));\n      if(!yOpt)return {ok:false,error:`وجدت قائمة السنة لكن لم أجد السنة ${year} داخلها.`};\n      yearInp.value=yOpt.value;\n    }else{\n      yearInp.value=String(year);\n    }\n"""
    if old_set not in text:
        raise SystemExit('Expected year assignment line was not found; refusing silent patch.')
    text = text.replace(old_set, new_set, 1)

    old_return = "    return {ok:true,year:dg(yearInp.value),month:dg(monthSel.options?.[monthSel.selectedIndex]?.textContent||''),category:nt(categorySel.options?.[categorySel.selectedIndex]?.textContent||'')};\n"
    new_return = """    const selectedYear=yearInp.tagName==='SELECT'?dg(`${yearInp.options?.[yearInp.selectedIndex]?.textContent||''} ${yearInp.value||''}`):dg(yearInp.value);\n    return {ok:true,year:selectedYear,month:dg(monthSel.options?.[monthSel.selectedIndex]?.textContent||''),category:nt(categorySel.options?.[categorySel.selectedIndex]?.textContent||'')};\n"""
    if old_return not in text:
        raise SystemExit('Expected prepareFiltersDirect return line was not found; refusing silent patch.')
    text = text.replace(old_return, new_return, 1)
    changed = True
    print('Applied robust V3 year-field detection')
else:
    print('Robust V3 year-field detection already present')

# ---------------------------------------------------------------------------
# 2) Do not configure filters until the disbursement screen is REALLY ready.
# ---------------------------------------------------------------------------
NAV_MARKER = '// DISBURSEMENT_READY_V1'
if NAV_MARKER not in text:
    old_nav = """async function goDisbursement(page){\n  if(page.url().includes(DISB_URL))return; log('➡️ الدخول إلى شاشة الصرفية...'); const btn=await clickable(page,'شاشة الصرفية'); if(!btn)throw new Error('لم أجد خيار \"شاشة الصرفية\".');\n  await btn.click({timeout:FAIL_TIMEOUT}); await waitUntil(async()=>page.url().includes('Family_sub2.aspx')||await bodyHas(page,['شاشة الصرفيات','شاشه الصرفيات']),'ظهور شاشة الصرفية'); log('✅ تم فتح شاشة الصرفية.');\n}\n"""

    new_nav = """// DISBURSEMENT_READY_V1\nasync function disbursementReady(page){\n  return page.evaluate(()=>{\n    const nt=v=>String(v||'').trim().replace(/[أإآ]/g,'ا').replace(/ى/g,'ي').replace(/ة/g,'ه').replace(/ـ/g,'').replace(/\\s+/g,' ');\n    const dg=v=>String(v||'').replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/\\D/g,'');\n    const vis=e=>{const s=getComputedStyle(e),r=e.getBoundingClientRect();return s.display!=='none'&&s.visibility!=='hidden'&&r.width>0&&r.height>0;};\n    const pathOk=String(location.pathname||'').toLowerCase().endsWith('/iccs/family_sub2.aspx');\n    const clickables=Array.from(document.querySelectorAll('button,input[type=\"submit\"],input[type=\"button\"],a'));\n    const hasShow=clickables.some(e=>vis(e)&&nt(`${e.value||''} ${e.innerText||e.textContent||''}`)==='عرض');\n    let hasMonth=false,hasCategory=false;\n    for(const s of document.querySelectorAll('select')){\n      if(!vis(s))continue;\n      const texts=Array.from(s.options||[]).map(o=>nt(o.textContent||''));\n      const vals=new Set(texts.map(dg).filter(Boolean));\n      if(Array.from({length:12},(_,i)=>String(i+1)).every(m=>vals.has(m)))hasMonth=true;\n      if(texts.some(x=>x.includes('ايتام'))&&texts.some(x=>x.includes('اسر'))&&texts.some(x=>x.includes('طلاب علم')))hasCategory=true;\n    }\n    return pathOk&&hasShow&&hasMonth&&hasCategory;\n  }).catch(()=>false);\n}\n\nasync function goDisbursement(page){\n  log('➡️ الدخول إلى شاشة الصرفية...');\n  if(await disbursementReady(page)){ log('✅ شاشة الصرفية جاهزة بالفعل.'); return; }\n  const btn=await clickable(page,'شاشة الصرفية');\n  if(!btn)throw new Error('لم أجد خيار \"شاشة الصرفية\" في الصفحة الرئيسية.');\n  await btn.click({timeout:FAIL_TIMEOUT});\n  await waitUntil(async()=>await disbursementReady(page),'اكتمال فتح شاشة الصرفية وظهور زر عرض والقوائم');\n  log('✅ تم فتح شاشة الصرفية واكتملت عناصرها.');\n}\n"""

    if old_nav not in text:
        raise SystemExit('Expected goDisbursement block was not found; refusing silent navigation patch.')
    text = text.replace(old_nav, new_nav, 1)
    changed = True
    print('Applied strict disbursement-screen readiness check')
else:
    print('Strict disbursement-screen readiness check already present')

if YEAR_MARKER not in text or NAV_MARKER not in text:
    raise SystemExit('Required browser compatibility markers are missing after patch.')

if changed:
    PATH.write_text(text, encoding='utf-8')
else:
    print('No browser compatibility changes needed')
