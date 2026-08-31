# -*- coding: utf-8 -*-
import os
import re
from difflib import SequenceMatcher
from decimal import Decimal

import pandas as pd
import tkinter as tk
from tkinter import messagebox

import kafala_compare_app as base


TOTAL_TOLERANCE = Decimal('0.0001')


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


def detail_total_including_duplicates(ref, ref_dup):
    """مجموع كل السجلات الفعلية، بما فيها الصفوف المكررة قبل حذف التكرار."""
    total = decimal_sum(ref['amount_ref'])
    if ref_dup is None or ref_dup.empty:
        return total
    # ref يحتوي بالفعل أول ظهور لكل رقم مكرر؛ نضيف فقط النسخ الزائدة.
    first_of_each_duplicate = ref_dup.drop_duplicates(subset=['nat_id'], keep='first')
    total += decimal_sum(ref_dup['amount_ref']) - decimal_sum(first_of_each_duplicate['amount_ref'])
    return total


def _total_label_rank(value):
    text = base.normalize_text(value)
    if not text:
        return 0
    # لا نريد مجاميع الصفحات؛ نبحث فقط عن إجمالي التقرير الكامل.
    if 'الصفحه' in text or 'الصفحة' in text:
        return 0
    patterns = [
        ('المجموع الاجمالي', 100),
        ('الاجمالي العام', 100),
        ('المجموع الكلي', 95),
        ('المجموع النهائي', 95),
        ('اجمالي التقرير', 95),
        ('اجمالي الكفالات', 90),
    ]
    for phrase, rank in patterns:
        if phrase in text:
            return rank
    return 0


def find_printed_total_in_raw(raw):
    """يستخرج الإجمالي المطبوع من ورقة خام، مع تفضيل أقرب رقم لعبارة الإجمالي."""
    if raw is None or raw.empty:
        return None

    matches = []
    rows, cols = raw.shape
    for r in range(rows):
        for c in range(cols):
            rank = _total_label_rank(raw.iat[r, c])
            if not rank:
                continue

            numeric = []
            # غالباً قيمة الإجمالي في نفس السطر؛ نسمح أيضاً بسطر ملاصق بسبب الخلايا المدمجة.
            for rr in (r, r + 1, r - 1):
                if rr < 0 or rr >= rows:
                    continue
                for cc in range(max(0, c - 8), min(cols, c + 9)):
                    if rr == r and cc == c:
                        continue
                    amount = base.parse_amount(raw.iat[rr, cc])
                    if amount is None:
                        continue
                    distance = abs(cc - c) + (abs(rr - r) * 4)
                    numeric.append((distance, cc, amount))

            if numeric:
                numeric.sort(key=lambda x: (x[0], x[1]))
                distance, _, amount = numeric[0]
                matches.append((rank, r, -distance, amount))

    if not matches:
        return None

    # نفضّل العبارة الأقوى، ثم آخر ظهور لها في التقرير، ثم الرقم الأقرب.
    matches.sort(key=lambda x: (x[0], x[1], x[2]), reverse=True)
    return Decimal(matches[0][3])


def read_printed_reference_total(ref_path):
    """يقرأ الإجمالي المطبوع من جميع أوراق ملف برنامج الرعاية إن وُجد."""
    try:
        wb = pd.ExcelFile(ref_path)
    except Exception:
        return None

    candidates = []
    for sheet_name in wb.sheet_names:
        try:
            raw = pd.read_excel(ref_path, sheet_name=sheet_name, header=None, dtype=str)
            value = find_printed_total_in_raw(raw)
            if value is not None:
                candidates.append(value)
        except Exception:
            continue

    if not candidates:
        return None
    # إذا تكرر الإجمالي نفسه في أكثر من ورقة نأخذه مرة واحدة؛ وإذا اختلف نأخذ آخر ورقة.
    return Decimal(candidates[-1])


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

    # فحص إجمالي برنامج الرعاية قبل أي حسابات تخص كرامة.
    detail_total = detail_total_including_duplicates(ref, ref_dup)
    printed_total = read_printed_reference_total(ref_path)
    if printed_total is None:
        total_guard = 'غير متاح - لم يتم العثور على إجمالي مطبوع في التقرير'
        total_difference = None
        official_care_total = detail_total
    else:
        total_difference = printed_total - detail_total
        if abs(total_difference) <= TOTAL_TOLERANCE:
            total_guard = 'نعم'
        else:
            total_guard = 'لا - يوجد فرق بين إجمالي التقرير ومجموع السجلات'
        official_care_total = printed_total

    duplicate_nat_ids = set(ref_dup.get('nat_id', pd.Series(dtype=str)).astype(str)) | set(site_dup.get('nat_id', pd.Series(dtype=str)).astype(str))
    ref_work = ref[~ref['nat_id'].astype(str).isin(duplicate_nat_ids)].copy()
    site_work = site[~site['nat_id'].astype(str).isin(duplicate_nat_ids)].copy()

    ref_ids = set(ref_work['nat_id'].astype(str))
    site_ids = set(site_work['nat_id'].astype(str))
    unmatched_site = site_work[~site_work['nat_id'].astype(str).isin(ref_ids)].copy()
    unmatched_ref = ref_work[~ref_work['nat_id'].astype(str).isin(site_ids)].copy()

    review_candidates, suspect_site_ids, suspect_ref_ids = build_review_candidates(unmatched_site, unmatched_ref)

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
    totals_reliable = len(ref_dup) == 0 and len(site_dup) == 0 and not total_guard.startswith('لا')

    summary_rows = [
        # نبقي هذا المفتاح للتوافق مع الأتمتة، لكنه يعكس الإجمالي المطبوع إن كان موجوداً.
        ('مجموع_برنامج_الرعاية', fmt_amount(official_care_total)),
        ('مجموع_برنامج_الرعاية_من_السجلات', fmt_amount(detail_total)),
        ('الإجمالي_المطبوع_في_تقرير_الرعاية', fmt_amount(printed_total) if printed_total is not None else 'غير متاح'),
        ('فرق_إجمالي_الرعاية', fmt_amount(total_difference) if total_difference is not None else 'غير متاح'),
        ('سلامة_إجمالي_الرعاية', total_guard),
        ('مجموع_كرامة_قبل_التعديل', fmt_amount(site_total_before)),
        ('المجموع_المتوقع_بعد_التعديل_الآلي', fmt_amount(expected_after_auto)),
        ('مطابقة_المجموع_موثوقة', 'نعم' if totals_reliable else 'لا - توجد ملاحظة تحتاج مراجعة'),
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
        'care_total': fmt_amount(official_care_total),
        'care_detail_total': fmt_amount(detail_total),
        'care_printed_total': fmt_amount(printed_total) if printed_total is not None else '',
        'care_total_difference': fmt_amount(total_difference) if total_difference is not None else '',
        'care_total_guard': total_guard,
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

            if s['care_total_guard'].startswith('لا'):
                total_note = (
                    '\n❌ تحذير إجمالي برنامج الرعاية:\n'
                    f"الإجمالي المطبوع في التقرير: {s['care_printed_total']}\n"
                    f"مجموع السجلات التي قرأها البرنامج: {s['care_detail_total']}\n"
                    f"الفرق: {s['care_total_difference']}\n"
                    'تم حجب التعديل التلقائي حتى يتم حل هذا الفرق.\n'
                )
            elif s['care_total_guard'].startswith('غير متاح'):
                total_note = (
                    '\n⚠ لم يتم العثور على إجمالي مطبوع في تقرير برنامج الرعاية؛ '
                    'تمت المقارنة لكن فحص الإجمالي غير متاح.\n'
                )
            else:
                total_note = (
                    '\n✅ إجمالي برنامج الرعاية سليم: '
                    f"{s['care_printed_total']} ومتطابق مع مجموع السجلات.\n"
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
                f"{review_note}{total_note}\n"
                f"المجموع المتوقع بعد الأتمتة الحالية: {s['expected_after_auto']}\n\n"
                f"ملف النتيجة:\n{out}"
            )

            if s['care_total_guard'].startswith('لا'):
                self.status.config(text='تمت المقارنة - التعديل التلقائي محجوب بسبب فرق الإجمالي')
                messagebox.showwarning('تحذير إجمالي برنامج الرعاية', msg)
            else:
                self.status.config(text='تمت العملية بنجاح')
                messagebox.showinfo('نجاح', msg)
        except Exception as e:
            self.status.config(text='حدث خطأ')
            messagebox.showerror('خطأ', str(e))

    def run_auto_update(self):
        """يمنع فتح الأتمتة إذا فشل فحص إجمالي برنامج الرعاية."""
        out_file = self.out_var.get().strip()
        if not out_file or not os.path.exists(out_file):
            messagebox.showerror('خطأ', 'نفّذ المقارنة أولاً باستخدام النسخة الحالية من البرنامج.')
            return

        try:
            summary_df = pd.read_excel(out_file, sheet_name='ملخص', dtype=str).fillna('')
            values = {
                str(row.iloc[0]).strip(): str(row.iloc[1]).strip()
                for _, row in summary_df.iterrows()
                if len(row) >= 2 and str(row.iloc[0]).strip()
            }
        except Exception:
            messagebox.showerror(
                'إعادة المقارنة مطلوبة',
                'ملف النتيجة لا يحتوي فحص إجمالي برنامج الرعاية.\n'
                'أعد تنفيذ المقارنة بهذه النسخة قبل بدء التعديل التلقائي.'
            )
            return

        guard = values.get('سلامة_إجمالي_الرعاية', '')
        if not guard:
            messagebox.showerror(
                'إعادة المقارنة مطلوبة',
                'لم أجد نتيجة فحص إجمالي برنامج الرعاية في ملف المقارنة.\n'
                'أعد المقارنة قبل المتابعة.'
            )
            return

        if guard.startswith('لا'):
            printed = values.get('الإجمالي_المطبوع_في_تقرير_الرعاية', 'غير متاح')
            detail = values.get('مجموع_برنامج_الرعاية_من_السجلات', 'غير متاح')
            diff = values.get('فرق_إجمالي_الرعاية', 'غير متاح')
            messagebox.showerror(
                'تم منع التعديل التلقائي للحماية',
                'يوجد اختلاف داخل ملف برنامج الرعاية نفسه:\n\n'
                f'الإجمالي المطبوع: {printed}\n'
                f'مجموع السجلات: {detail}\n'
                f'الفرق: {diff}\n\n'
                'لن يتم فتح موقع كرامة أو تنفيذ أي تعديل حتى يتم حل الفرق.'
            )
            return

        if guard.startswith('غير متاح'):
            if not messagebox.askyesno(
                'تعذر فحص إجمالي التقرير',
                'لم يتم العثور على إجمالي مطبوع داخل ملف برنامج الرعاية.\n'
                'يمكن المتابعة، لكن فحص مطابقة إجمالي التقرير غير متاح.\n\nهل تريد المتابعة؟'
            ):
                return

        super().run_auto_update()


if __name__ == '__main__':
    root = tk.Tk()
    app = EnhancedApp(root)
    root.mainloop()
