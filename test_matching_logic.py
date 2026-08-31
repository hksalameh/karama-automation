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
