# -*- coding: utf-8 -*-
import re
from difflib import SequenceMatcher
from decimal import Decimal
from datetime import datetime

import pandas as pd
import tkinter as tk
from tkinter import messagebox

import kafala_compare_app as base


def name_key(value):
    """مفتاح اسم للمراجعة فقط: يوحد الحروف ويحذف المسافات والرموز."""
    text = base.normalize_text(value)
    text = text.replace('ؤ', 'و').replace('ئ', 'ي')
    text = re.sub(r'[^\w\u0600-\u06FF]+', '', text, flags=re.UNICODE)
    return text


def name_similarity(a, b):
    ka, kb = name_key(a), name_key(b)
    if not ka or not kb:
        return 0.0
    return SequenceMatcher(None, ka, kb).ratio()


def nat_digit_diff(a, b):
    a, b = base.normalize_nat_id(a), base.normalize_nat_id(b)
    if not a or not b or len(a) != len(b):
        return None
    return sum(x != y for x, y in zip(a, b))


def fmt_amount(d):
    if d is None or pd.isna(d):
        return ''
    d = Decimal(d)
    d = d.quantize(Decimal('0.0001')).normalize()
    s = format(d, 'f')
    return s.rstrip('0').rstrip('.') if '.' in s else s


def decimal_sum(series):
    total = Decimal('0')
    for value in series:
        if value is not None and not pd.isna(value):
            total += Decimal(value)
    return total


def build_review_candidates(unmatched_site, unmatched_ref):
    rows = []
    suspect_site = set()
    suspect_ref = set()

    for _, s in unmatched_site.iterrows():
        skey = name_key(s.get('name_site', ''))
        if len(skey) < 5:
            continue
        candidates = []
        for _, r in unmatched_ref.iterrows():
            rkey = name_key(r.get('name_ref', ''))
            if len(rkey) < 5:
                continue
            sim = name_similarity(s.get('name_site', ''), r.get('name_ref', ''))
            diff = nat_digit_diff(s.get('nat_id', ''), r.get('nat_id', ''))
            exact_compact = skey == rkey
            qualifies = exact_compact or sim >= 0.94 or (diff is not None and diff <= 1 and sim >= 0.80)
            if not qualifies:
                continue
            score = sim + (0.10 if exact_compact else 0) + (0.05 if diff is not None and diff <= 1 else 0)
            candidates.append((score, sim, diff, exact_compact, r))

        if not candidates:
            continue

        candidates.sort(key=lambda x: x[0], reverse=True)
        # نحتفظ بأفضل الاحتمالات القريبة فقط حتى لا يصبح التقرير مزدحماً.
        best_score = candidates[0][0]
        selected = [c for c in candidates if c[0] >= best_score - 0.035][:5]
        ambiguous = len(selected) > 1

        for score, sim, diff, exact_compact, r in selected:
            suspect_site.add(str(s['nat_id']))
            suspect_ref.add(str(r['nat_id']))
            if ambiguous:
                status = 'اسم مكرر/أكثر من احتمال - مراجعة يدوية'
            elif exact_compact:
                status = 'الاسم متطابق بعد تجاهل المسافات/الرموز - الرقم الوطني مختلف'
            elif diff is not None and diff <= 1:
                status = 'احتمال خطأ برقم وطني مع تشابه قوي بالاسم'
            else:
                status = 'تشابه قوي بالاسم - مراجعة الرقم الوطني'

            rows.append({
                'رقم_وطني_كرامة': s['nat_id'],
                'الاسم_في_كرامة': s.get('name_site', ''),
                'مبلغ_كرامة': fmt_amount(s.get('amount_site')),
                'رقم_وطني_الرعاية': r['nat_id'],
                'الاسم_في_الرعاية': r.get('name_ref', ''),
                'مبلغ_الرعاية': fmt_amount(r.get('amount_ref')),
                'تشابه_الاسم_%': round(sim * 100, 1),
                'عدد_خانات_الرقم_المختلفة': '' if diff is None else diff,
                'الحالة': status,
                'التوصية': 'مراجعة يدوية فقط - ممنوع التصفير أو الإضافة الآلية لهذه الحالة',
            })

    return pd.DataFrame(rows), suspect_site, suspect_ref


def enhanced_compare_and_export(ref_path, site_path, output_path):
    ref, ref_dup = base.read_reference(ref_path)
    site, site_dup = base.read_site(site_path)

    duplicate_nat_ids = set(ref_dup.get('nat_id', pd.Series(dtype=str)).astype(str)) | set(site_dup.get('nat_id', pd.Series(dtype=str)).astype(str))
    ref_work = ref[~ref['nat_id'].astype(str).isin(duplicate_nat_ids)].copy()
    site_work = site[~site['nat_id'].astype(str).isin(duplicate_nat_ids)].copy()

    ref_ids = set(ref_work['nat_id'].astype(str))
    site_ids = set(site_work['nat_id'].astype(str))
    unmatched_site = site_work[~site_work['nat_id'].astype(str).isin(ref_ids)].copy()
    unmatched_ref = ref_work[~ref_work['nat_id'].astype(str).isin(site_ids)].copy()

    review_candidates, suspect_site_ids, suspect_ref_ids = build_review_candidates(unmatched_site, unmatched_ref)

    # الحالات المشتبه بها لا تدخل في أي تعديل آلي.
    safe_site = site_work[~site_work['nat_id'].astype(str).isin(suspect_site_ids)].copy()
    safe_ref = ref_work[~ref_work['nat_id'].astype(str).isin(suspect_ref_ids)].copy()

    merged = safe_site.merge(safe_ref, on='nat_id', how='left')
    missing_in_site_raw = safe_ref.merge(safe_site[['nat_id']], on='nat_id', how='left', indicator=True)
    missing_in_site_raw = missing_in_site_raw[missing_in_site_raw['_merge'] == 'left_only'].drop(columns=['_merge'])

    merged['amount_expected'] = merged['amount_ref'].apply(lambda x: x if pd.notna(x) else Decimal('0'))
    merged['exists_in_ref'] = merged['amount_ref'].notna()
    merged['is_match'] = merged.apply(lambda r: base.almost_equal(r['amount_site'], r['amount_expected']), axis=1)

    report = merged.copy()
    report['المبلغ_في_الموقع'] = report['amount_site'].map(fmt_amount)
    report['المبلغ_الفعلي'] = report['amount_expected'].map(fmt_amount)
    report['الرقم_الوطني'] = report['nat_id']
    report['الاسم_في_الموقع'] = report['name_site']
    report['الاسم_في_برنامج_الرعاية'] = report['name_ref'].fillna('غير موجود ببرنامج الرعاية')
    report['سبب'] = report.apply(
        lambda r: 'مطابق' if r['is_match'] else ('غير موجود ببرنامج الرعاية -> المتوقع 0' if not r['exists_in_ref'] else 'اختلاف مبلغ'), axis=1
    )
    report['تعليمات'] = report.apply(
        lambda r: '' if r['is_match'] else f"تعديل من {r['المبلغ_في_الموقع']} الى {r['المبلغ_الفعلي']}", axis=1
    )

    cols = ['الرقم_الوطني', 'الاسم_في_الموقع', 'الاسم_في_برنامج_الرعاية', 'المبلغ_في_الموقع', 'المبلغ_الفعلي', 'سبب', 'تعليمات']
    needs_update = report[~report['is_match']][cols].sort_values('الرقم_الوطني')
    matches = report[report['is_match']][cols].sort_values('الرقم_الوطني')
    missing_in_ref = report[(~report['exists_in_ref']) & (~report['is_match'])][cols].sort_values('الرقم_الوطني')
    amount_diff = report[(report['exists_in_ref']) & (~report['is_match'])][cols].sort_values('الرقم_الوطني')

    missing_in_site = pd.DataFrame({
        'الرقم_الوطني': missing_in_site_raw['nat_id'],
        'الاسم_في_الموقع': '',
        'الاسم_في_برنامج_الرعاية': missing_in_site_raw['name_ref'],
        'المبلغ_في_الموقع': '',
        'المبلغ_الفعلي': missing_in_site_raw['amount_ref'].map(fmt_amount),
        'سبب': 'موجود ببرنامج الرعاية وغير موجود بالموقع',
        'تعليمات': 'اضافة للموقع',
    })[cols].sort_values('الرقم_الوطني')

    # مجموع متوقع بعد هذه الأتمتة تحديداً: الحالات المشتبه بها تبقى بقيمها الحالية دون لمس.
    ref_map = {str(r['nat_id']): Decimal(r['amount_ref']) for _, r in ref_work.iterrows()}
    expected_after_auto = Decimal('0')
    for _, row in site_work.iterrows():
        nat = str(row['nat_id'])
        current = Decimal(row['amount_site'])
        if nat in suspect_site_ids or nat in duplicate_nat_ids:
            expected_after_auto += current
        elif nat in ref_map:
            expected_after_auto += ref_map[nat]
        else:
            expected_after_auto += Decimal('0')

    site_total_before = decimal_sum(site['amount_site'])
    care_total = decimal_sum(ref['amount_ref'])
    totals_reliable = len(ref_dup) == 0 and len(site_dup) == 0

    summary_rows = [
        ('مجموع_برنامج_الرعاية', fmt_amount(care_total)),
        ('مجموع_كرامة_قبل_التعديل', fmt_amount(site_total_before)),
        ('المجموع_المتوقع_بعد_التعديل_الآلي', fmt_amount(expected_after_auto)),
        ('مطابقة_المجموع_موثوقة', 'نعم' if totals_reliable else 'لا - يوجد رقم وطني مكرر'),
        ('حالات_مراجعة_مطابقة_محتملة', len(review_candidates)),
        ('عدد_الأرقام_الوطنية_المكررة_بالرعاية', len(ref_dup)),
        ('عدد_الأرقام_الوطنية_المكررة_بكرامة', len(site_dup)),
        ('غير_موجود_بكرامة_بعد_استبعاد_المشتبه', len(missing_in_site)),
        ('سيتم_تصفيرها_آلياً', len(missing_in_ref)),
        ('اختلاف_مبلغ_آلي', len(amount_diff)),
    ]
    summary_df = pd.DataFrame(summary_rows, columns=['البند', 'القيمة'])

    with pd.ExcelWriter(output_path, engine='openpyxl') as w:
        needs_update.to_excel(w, index=False, sheet_name='يحتاج تعديل')
        missing_in_ref.to_excel(w, index=False, sheet_name='غير موجود ببرنامج الرعاية')
        missing_in_site.to_excel(w, index=False, sheet_name='غير موجود بالموقع')
        amount_diff.to_excel(w, index=False, sheet_name='اختلاف مبلغ فقط')
        matches.to_excel(w, index=False, sheet_name='مطابق')
        summary_df.to_excel(w, index=False, sheet_name='ملخص')
        if not review_candidates.empty:
            review_candidates.to_excel(w, index=False, sheet_name='مراجعة مطابقة محتملة')
        if not ref_dup.empty:
            ref_dup.to_excel(w, index=False, sheet_name='تكرار ببرنامج الرعاية')
        if not site_dup.empty:
            site_dup.to_excel(w, index=False, sheet_name='تكرار بالموقع')

    summary = {
        'site_rows': len(site),
        'ref_rows': len(ref),
        'needs_update': len(needs_update),
        'matches': len(matches),
        'missing_in_ref_updates_to_zero': len(missing_in_ref),
        'missing_in_site_needs_add': len(missing_in_site),
        'amount_diff_only': len(amount_diff),
        'ref_duplicates': len(ref_dup),
        'site_duplicates': len(site_dup),
        'review_candidates': len(review_candidates),
        'care_total': fmt_amount(care_total),
        'expected_after_auto': fmt_amount(expected_after_auto),
    }
    return summary, needs_update, missing_in_ref, amount_diff, missing_in_site


base.compare_and_export = enhanced_compare_and_export


class EnhancedApp(base.App):
    def run_compare(self):
        ref = self.ref_var.get().strip()
        site = self.site_var.get().strip()
        out = self.out_var.get().strip()
        if not ref or not site:
            messagebox.showerror('خطأ', 'يرجى اختيار ملف المرجع وملف الموقع.')
            return
        if not out:
            self.suggest_output()
            out = self.out_var.get().strip()
        try:
            self.status.config(text='جاري المعالجة...')
            self.root.update_idletasks()
            s, needs_update, missing_df, amount_diff_df, missing_site_df = enhanced_compare_and_export(ref, site, out)
            self.needs_update_df = needs_update.reset_index(drop=True)
            self.missing_df = missing_df.reset_index(drop=True)
            self.amount_diff_df = amount_diff_df.reset_index(drop=True)
            self.missing_site_df = missing_site_df.reset_index(drop=True)
            self.current_idx = 0
            self.set_review_mode('all')
            review_note = ''
            if s['review_candidates']:
                review_note = (
                    f"\n⚠ حالات مطابقة محتملة تحتاج مراجعة يدوية: {s['review_candidates']}\n"
                    'تم استبعادها من التصفير والإضافة والتعديل الآلي.\n'
                    'راجع ورقة: مراجعة مطابقة محتملة\n'
                )
            msg = (
                f"تمت المقارنة بنجاح\n\n"
                f"سجلات كرامة: {s['site_rows']}\n"
                f"سجلات برنامج الرعاية: {s['ref_rows']}\n"
                f"يحتاج تعديل آمن: {s['needs_update']}\n"
                f"مطابق: {s['matches']}\n"
                f"سيتم التصفير: {s['missing_in_ref_updates_to_zero']}\n"
                f"موجود بالرعاية وغير موجود بكرامة: {s['missing_in_site_needs_add']}\n"
                f"اختلاف مبلغ فقط: {s['amount_diff_only']}\n"
                f"تكرار رقم وطني بالرعاية: {s['ref_duplicates']}\n"
                f"تكرار رقم وطني بكرامة: {s['site_duplicates']}\n"
                f"{review_note}\n"
                f"مجموع برنامج الرعاية: {s['care_total']}\n"
                f"المجموع المتوقع بعد الأتمتة الحالية: {s['expected_after_auto']}\n\n"
                f"ملف النتيجة:\n{out}"
            )
            self.status.config(text='تمت العملية بنجاح')
            messagebox.showinfo('نجاح', msg)
        except Exception as e:
            self.status.config(text='حدث خطأ')
            messagebox.showerror('خطأ', str(e))


if __name__ == '__main__':
    root = tk.Tk()
    app = EnhancedApp(root)
    root.mainloop()
