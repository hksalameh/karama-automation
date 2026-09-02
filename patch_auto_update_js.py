# -*- coding: utf-8 -*-
"""Apply small, idempotent safety fixes to the bundled Karama browser script before build."""
from pathlib import Path

PATH = Path(__file__).with_name('auto_update_from_diff.js')
text = PATH.read_text(encoding='utf-8')

MARKER = '// YEAR_FIELD_ROBUST_V2'
if MARKER in text:
    print('auto_update_from_diff.js already has robust year-field detection')
    raise SystemExit(0)

old_prepare = """    for(const i of document.querySelectorAll('input')){\n      if(!vis(i)||/hidden|submit|button|password|checkbox|radio|file/.test(String(i.type||'text').toLowerCase()))continue;\n      const val=dg(i.value),ctx=nt(i.closest('tr')?.innerText||i.parentElement?.innerText||'');\n      if((/^\\d{4}$/.test(val)||ctx.includes('سنه')||ctx.includes('سنة'))&&(ctx.includes('سنه')||ctx.includes('سنة'))){yearInp=i;break;}\n    }\n"""

new_prepare = """    // YEAR_FIELD_ROBUST_V2\n    // موقع كرامة لا يضع تسمية السنة دائماً داخل نفس <tr> مع الحقل.\n    // لذلك نعتمد أولاً على الحقول المرئية ذات الأربع خانات، ثم نستخدم القرب البصري من كلمة السنة عند الحاجة.\n    const usableInputs=Array.from(document.querySelectorAll('input')).filter(i=>\n      vis(i)&&!/hidden|submit|button|password|checkbox|radio|file/.test(String(i.type||'text').toLowerCase())\n    );\n    const fourDigit=usableInputs.filter(i=>/^\\d{4}$/.test(dg(i.value)));\n    if(fourDigit.length===1){\n      yearInp=fourDigit[0];\n    }else if(fourDigit.length>1){\n      yearInp=fourDigit.find(i=>dg(i.value)===String(year))||null;\n      if(!yearInp){\n        const yearLabels=Array.from(document.querySelectorAll('label,td,th,span,div')).filter(e=>{\n          if(!vis(e))return false;\n          const t=nt(e.innerText||e.textContent||'');\n          return t==='السنه'||t==='السنة'||t.includes('السنه')||t.includes('السنة');\n        });\n        let best=null,bestDistance=Infinity;\n        for(const i of fourDigit){\n          const ir=i.getBoundingClientRect();\n          for(const l of yearLabels){\n            const lr=l.getBoundingClientRect();\n            const d=Math.abs((ir.top+ir.height/2)-(lr.top+lr.height/2))*5+Math.abs((ir.left+ir.width/2)-(lr.left+lr.width/2));\n            if(d<bestDistance){bestDistance=d;best=i;}\n          }\n        }\n        yearInp=best;\n      }\n    }\n    if(!yearInp){\n      yearInp=usableInputs.find(i=>/year|sana|sanah|سنه|سنة/i.test(`${i.id||''} ${i.name||''} ${i.getAttribute('aria-label')||''}`))||null;\n    }\n"""

old_state = """    for(const i of document.querySelectorAll('input')){ if(!vis(i)||/hidden|submit|button|password|checkbox|radio|file/.test(String(i.type||'text')))continue; const v=dg(i.value),ctx=nt(i.closest('tr')?.innerText||i.parentElement?.innerText||''); if(/^\\d{4}$/.test(v)&&(ctx.includes('سنه')||ctx.includes('سنة'))){year=v;break;} }\n"""

new_state = """    for(const i of document.querySelectorAll('input')){ if(!vis(i)||/hidden|submit|button|password|checkbox|radio|file/.test(String(i.type||'text')))continue; const v=dg(i.value); if(/^\\d{4}$/.test(v)){year=v;break;} }\n"""

if old_prepare not in text:
    raise SystemExit('Expected prepareFiltersDirect year-detection block was not found; refusing silent patch.')
if old_state not in text:
    raise SystemExit('Expected screenState year-detection block was not found; refusing silent patch.')

text = text.replace(old_prepare, new_prepare, 1)
text = text.replace(old_state, new_state, 1)
PATH.write_text(text, encoding='utf-8')
print('Applied robust year-field detection to auto_update_from_diff.js')
