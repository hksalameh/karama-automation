from decimal import Decimal
import pandas as pd

from kafala_compare_app_v2 import (
    build_review_candidates,
    detail_total_including_duplicates,
    find_printed_total_in_raw,
)


def row_site(nat, name, amount='24'):
    return {'nat_id': nat, 'name_site': name, 'amount_site': Decimal(amount)}


def row_ref(nat, name, amount='24'):
    return {'nat_id': nat, 'name_ref': name, 'amount_ref': Decimal(amount)}


# عبدالله / عبد الله مع رقم مختلف: يجب حمايته من التصفير.
site = pd.DataFrame([row_site('2001234567', 'محمد عبدالله احمد')])
ref = pd.DataFrame([row_ref('2001234568', 'محمد عبد الله احمد')])
review, s_ids, r_ids = build_review_candidates(site, ref)
assert not review.empty
assert '2001234567' in s_ids and '2001234568' in r_ids

# خطأ حرف بالاسم + خانة واحدة بالرقم الوطني: مراجعة فقط.
site = pd.DataFrame([row_site('2012345678', 'احمد محمد الزعبي')])
ref = pd.DataFrame([row_ref('2012345679', 'احمد محمذ الزعبي')])
review, s_ids, r_ids = build_review_candidates(site, ref)
assert not review.empty
assert '2012345678' in s_ids and '2012345679' in r_ids

# شخص مختلف فعلاً: لا يجب اختراع مطابقة.
site = pd.DataFrame([row_site('3000000001', 'سليم خالد العلي')])
ref = pd.DataFrame([row_ref('4000000002', 'نور محمود الخطيب')])
review, s_ids, r_ids = build_review_candidates(site, ref)
assert review.empty and not s_ids and not r_ids

# اسم مكرر في الرعاية: يجب إظهار أكثر من احتمال وعدم الحسم آلياً.
site = pd.DataFrame([row_site('5000000001', 'محمد احمد علي')])
ref = pd.DataFrame([
    row_ref('5000000011', 'محمد احمد علي'),
    row_ref('5000000022', 'محمد احمد علي'),
])
review, s_ids, r_ids = build_review_candidates(site, ref)
assert len(review) == 2
assert review['الحالة'].str.contains('مكرر').all()

# يجب تجاهل مجموع الصفحة واختيار المجموع الإجمالي المطبوع.
raw = pd.DataFrame([
    ['', 'مجموع الصفحة', '412', ''],
    ['', '', '', ''],
    ['', 'المجموع الإجمالي', '7633.3', ''],
])
assert find_printed_total_in_raw(raw) == Decimal('7633.3')

# مجموع التفاصيل يجب أن يشمل الصفوف المكررة كلها، لا أول سجل فقط.
ref_unique = pd.DataFrame([
    row_ref('1', 'أ', '10'),
    row_ref('2', 'ب', '20'),
])
ref_dup = pd.DataFrame([
    row_ref('2', 'ب', '20'),
    row_ref('2', 'ب', '4'),
])
assert detail_total_including_duplicates(ref_unique, ref_dup) == Decimal('34')

print('matching and total safety tests passed')

# Review data must stay accessible in-app without changing the workbook or IDs.
import os
import tempfile
from pathlib import Path
from openpyxl import Workbook
from auto_update_gui import read_review_cases, ReviewCasesWindow

with tempfile.TemporaryDirectory() as folder:
    path = Path(folder) / 'review.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = 'مراجعة مطابقة محتملة'
    ws.append(['رقم_وطني_كرامة', 'الاسم_في_كرامة', 'رقم_وطني_الرعاية', 'الحالة'])
    ws.append(['0012345678', 'محمد عبدالله احمد', '0012345679', 'اختلاف الرقم الوطني'])
    ws.append(['0098765432', 'نور خالد احمد', '0098765433', 'تشابه الاسم'])
    ws.append([None, None, None, None])
    dup = wb.create_sheet('تكرار ببرنامج الرعاية')
    dup.append(['nat_id', 'name_ref', 'amount_ref'])
    dup.append(['0012345679', 'محمد عبدالله احمد', 24])
    summary = wb.create_sheet('ملخص')
    summary.append(['البند', 'القيمة'])
    summary.append(['سلامة_إجمالي_الرعاية', 'لا - يوجد فرق'])
    wb.save(path)
    original = path.read_bytes()
    cases = read_review_cases(path)
    assert len(cases['مراجعة مطابقة محتملة']) == 2
    assert cases['مراجعة مطابقة محتملة'][0]['رقم_وطني_كرامة'] == '0012345678'
    assert cases['تكرار ببرنامج الرعاية'][0]['nat_id'] == '0012345679'
    assert cases['ملخص'][0]['القيمة'] == 'لا - يوجد فرق'
    assert path.read_bytes() == original

    # The Windows build runner can exercise real Tk widgets and clipboard.
    if os.name == 'nt' or os.environ.get('DISPLAY'):
        import tkinter as tk
        from kafala_compare_app_v2 import EnhancedApp
        from auto_update_gui import AutoUpdateGUI
        root = tk.Tk()
        root.withdraw()
        try:
            app = EnhancedApp(root)
            window = ReviewCasesWindow(root, cases)
            root.update_idletasks()
            assert len(window.tree.get_children()) == 2
            window.move(1)
            assert window.tree.selection() == ('1',)
            window.copy_number('site')
            assert root.clipboard_get() == '0098765432'
            window.copy_number('care')
            assert root.clipboard_get() == '0098765433'
            window.query.set('محمد')
            assert len(window.tree.get_children()) == 1
            assert '0012345678' in window.details.get('1.0', tk.END)
            window.query.set('لا يوجد هذا الاسم')
            assert len(window.tree.get_children()) == 0
            window.move(1)
            window.query.set('')
            window.category.set('تكرار ببرنامج الرعاية')
            window.refresh()
            window.copy_number('care')
            assert root.clipboard_get() == '0012345679'
            panel = tk.Frame(root)
            automation = AutoUpdateGUI(panel, str(path), {})
            window.win.destroy()
        finally:
            root.destroy()
    assert path.read_bytes() == original
print('in-app review data tests passed')
