import os
import re
import sys
import time
import subprocess
from decimal import Decimal, InvalidOperation
from datetime import datetime
import tkinter as tk
from tkinter import font as tkfont
from tkinter import filedialog, messagebox
from tkinter import ttk
import tkinter.scrolledtext

import pandas as pd


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Not in a PyInstaller bundle
        base_path = os.path.abspath(os.path.dirname(__file__))
        if os.path.basename(base_path).lower() == '__pycache__':
            base_path = os.path.abspath(os.path.join(base_path, os.pardir))

    return os.path.join(base_path, relative_path)



def normalize_text(s):
    if pd.isna(s):
        return ''
    s = str(s).strip()
    # normalize Arabic variants
    s = s.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
    s = s.replace('ى', 'ي')
    s = s.replace('ـ', '')
    s = re.sub(r'\s+', ' ', s)
    return s


def normalize_nat_id(v):
    if pd.isna(v):
        return ''
    s = str(v).strip()
    if s.endswith('.0'):
        s = s[:-2]
    # arabic-indic digits to latin
    trans = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')
    s = s.translate(trans)
    s = re.sub(r'\D+', '', s)
    return s


def parse_amount(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    if not s:
        return None
    trans = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')
    s = s.translate(trans).replace(',', '')
    s = re.sub(r'[^0-9.\-]', '', s)
    if s in ('', '.', '-', '-.'):
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def almost_equal(a, b, tol=Decimal('0.0001')):
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


def detect_columns(df):
    cols = list(df.columns)
    norm = {c: normalize_text(c) for c in cols}

    def pick(parts):
        parts = [normalize_text(p) for p in parts]
        for p in parts:
            for c, nc in norm.items():
                if p in nc:
                    return c
        return None

    nat_col = pick(['الرقم الوطني', 'رقم الوطني', 'الوطني'])
    name_col = pick(['اسم اليتيم', 'الاسم عربي', 'الاسم', 'اسم'])
    amount_col = pick(['قيمة الكفالة', 'مبلغ الكفالة', 'المبلغ', 'قيمة'])
    return nat_col, name_col, amount_col


def pick_amount_column_near(raw, header_row, amount_header_col):
    candidates = []
    for offset in (0, -1, 1, -2, 2):
        c = amount_header_col + offset
        if 0 <= c < raw.shape[1] and c not in candidates:
            candidates.append(c)

    best_col = amount_header_col
    best_score = -1
    data = raw.iloc[header_row + 1:]
    for c in candidates:
        score = data.iloc[:, c].map(parse_amount).notna().sum()
        if score > best_score:
            best_score = score
            best_col = c
    return best_col


def pick_name_column_near(raw, header_row, name_header_col):
    candidates = []
    for offset in (0, -1, 1, -2, 2):
        c = name_header_col + offset
        if 0 <= c < raw.shape[1] and c not in candidates:
            candidates.append(c)

    best_col = name_header_col
    best_score = -1
    data = raw.iloc[header_row + 1:]
    for c in candidates:
        score = 0
        for value in data.iloc[:, c]:
            text = normalize_text(value)
            if text and len(text) >= 5 and not normalize_nat_id(text) and parse_amount(text) is None:
                score += 1
        if score > best_score:
            best_score = score
            best_col = c
    return best_col


def read_report_layout(raw):
    max_header_scan = min(len(raw), 80)
    for header_row in range(max_header_scan):
        labels = {c: normalize_text(v) for c, v in raw.iloc[header_row].items() if pd.notna(v)}
        nat_col = next((c for c, v in labels.items() if 'الرقم الوطني' in v or 'رقم وطني' in v or 'الوطني' in v), None)
        name_col = next((c for c, v in labels.items() if 'اسم اليتيم' in v or 'اسم المعيل' in v or 'الاسم' in v or 'اسم' in v), None)
        amount_header_col = next((c for c, v in labels.items() if 'قيمة الكفالة' in v or 'مبلغ الكفالة' in v or 'المبلغ' in v or 'قيمة' in v), None)

        if nat_col is None or amount_header_col is None:
            continue

        amount_col = pick_amount_column_near(raw, header_row, amount_header_col)
        if name_col is not None:
            name_col = pick_name_column_near(raw, header_row, name_col)
        data = raw.iloc[header_row + 1:].copy()
        p = pd.DataFrame({
            'name_ref': data.iloc[:, name_col].map(normalize_text) if name_col is not None else '',
            'nat_id': data.iloc[:, nat_col].map(normalize_nat_id),
            'amount_ref': data.iloc[:, amount_col].map(parse_amount),
        })
        p = p[(p['nat_id'] != '') & (p['amount_ref'].notna())]
        if not p.empty:
            return p
    return pd.DataFrame()


def read_reference(ref_path):
    wb = pd.ExcelFile(ref_path)
    parts = []

    for sheet in wb.sheet_names:
        # 1) Try header-based detection
        try:
            df = pd.read_excel(ref_path, sheet_name=sheet, dtype=str)
            nat_col, name_col, amount_col = detect_columns(df)
            if nat_col and amount_col:
                p = pd.DataFrame({
                    'name_ref': df[name_col].map(normalize_text) if name_col else '',
                    'nat_id': df[nat_col].map(normalize_nat_id),
                    'amount_ref': df[amount_col].map(parse_amount),
                })
                p = p[(p['nat_id'] != '') & (p['amount_ref'].notna())]
                if not p.empty:
                    parts.append(p)
                    continue
        except Exception:
            pass

        # 2) Try report layout with title rows and shifted/merged amount column
        try:
            raw = pd.read_excel(ref_path, sheet_name=sheet, header=None, dtype=str)
            p = read_report_layout(raw)
            if not p.empty:
                parts.append(p)
                continue
        except Exception:
            pass

        # 3) Try compact structure (name, nat_id, amount) in first 3 columns
        try:
            raw = pd.read_excel(ref_path, sheet_name=sheet, header=None, dtype=str)
            if raw.shape[1] >= 3:
                p = pd.DataFrame({
                    'name_ref': raw.iloc[:, 0].map(normalize_text),
                    'nat_id': raw.iloc[:, 1].map(normalize_nat_id),
                    'amount_ref': raw.iloc[:, 2].map(parse_amount),
                })
                p = p[(p['nat_id'] != '') & (p['amount_ref'].notna())]
                if not p.empty:
                    parts.append(p)
                    continue
        except Exception:
            pass

        # 4) Try detailed positional structure used by some old reports
        # name@20, nat@15, amount@10/11
        try:
            raw = pd.read_excel(ref_path, sheet_name=sheet, header=None, dtype=str)
            if raw.shape[1] > 20:
                amount_col = 10 if raw.iloc[:, 10].map(parse_amount).notna().sum() > raw.iloc[:, 11].map(parse_amount).notna().sum() else 11
                p = pd.DataFrame({
                    'name_ref': raw.iloc[:, 20].map(normalize_text),
                    'nat_id': raw.iloc[:, 15].map(normalize_nat_id),
                    'amount_ref': raw.iloc[:, amount_col].map(parse_amount),
                })
                p = p[(p['nat_id'] != '') & (p['amount_ref'].notna())]
                if not p.empty:
                    parts.append(p)
                    continue
        except Exception:
            pass

    if not parts:
        raise ValueError('تعذر قراءة ملف المرجع بصيغة مفهومة (اسم/رقم وطني/مبلغ).')

    ref = pd.concat(parts, ignore_index=True)
    ref = ref.dropna(subset=['nat_id', 'amount_ref'])
    ref = ref[ref['nat_id'] != '']

    duplicates = ref[ref.duplicated(subset=['nat_id'], keep=False)].copy()
    ref = ref.drop_duplicates(subset=['nat_id'], keep='first')
    return ref, duplicates


def read_site(site_path):
    site = pd.read_excel(site_path, dtype=str)
    nat_col, name_col, amount_col = detect_columns(site)

    if not nat_col or not amount_col:
        raise ValueError('تعذر اكتشاف أعمدة الرقم الوطني/المبلغ في ملف الموقع.')

    out = pd.DataFrame({
        'nat_id': site[nat_col].map(normalize_nat_id),
        'name_site': site[name_col].map(normalize_text) if name_col else '',
        'amount_site': site[amount_col].map(parse_amount),
    })

    out = out[(out['nat_id'] != '') & (out['amount_site'].notna())]
    duplicates = out[out.duplicated(subset=['nat_id'], keep=False)].copy()
    out = out.drop_duplicates(subset=['nat_id'], keep='first')
    return out, duplicates


def compare_and_export(ref_path, site_path, output_path):
    ref, ref_dup = read_reference(ref_path)
    site, site_dup = read_site(site_path)

    merged = site.merge(ref, on='nat_id', how='left')
    missing_in_site_raw = ref.merge(site[['nat_id']], on='nat_id', how='left', indicator=True)
    missing_in_site_raw = missing_in_site_raw[missing_in_site_raw['_merge'] == 'left_only'].drop(columns=['_merge'])

    # business rule: if missing in reference => expected amount is 0
    merged['amount_expected'] = merged['amount_ref'].apply(lambda x: x if pd.notna(x) else Decimal('0'))
    merged['exists_in_ref'] = merged['amount_ref'].notna()

    merged['is_match'] = merged.apply(
        lambda r: almost_equal(r['amount_site'], r['amount_expected']), axis=1
    )

    def fmt(d):
        if d is None or pd.isna(d):
            return ''
        d = Decimal(d)
        d = d.quantize(Decimal('0.0001')).normalize()
        return format(d, 'f').rstrip('0').rstrip('.') if '.' in format(d, 'f') else format(d, 'f')

    report = merged.copy()
    report['المبلغ_في_الموقع'] = report['amount_site'].map(fmt)
    report['المبلغ_الفعلي'] = report['amount_expected'].map(fmt)
    report['الرقم_الوطني'] = report['nat_id']
    report['الاسم_في_الموقع'] = report['name_site']
    report['الاسم_في_برنامج_الرعاية'] = report['name_ref'].fillna('غير موجود ببرنامج الرعاية')

    report['سبب'] = report.apply(
        lambda r: 'مطابق' if r['is_match'] else ('غير موجود ببرنامج الرعاية -> المتوقع 0' if not r['exists_in_ref'] else 'اختلاف مبلغ'),
        axis=1
    )
    report['تعليمات'] = report.apply(
        lambda r: '' if r['is_match'] else f"تعديل من {r['المبلغ_في_الموقع']} الى {r['المبلغ_الفعلي']}",
        axis=1
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
        'المبلغ_الفعلي': missing_in_site_raw['amount_ref'].map(fmt),
        'سبب': 'موجود ببرنامج الرعاية وغير موجود بالموقع',
        'تعليمات': 'اضافة للموقع',
    })[cols].sort_values('الرقم_الوطني')

    with pd.ExcelWriter(output_path, engine='openpyxl') as w:
        needs_update.to_excel(w, index=False, sheet_name='يحتاج تعديل')
        missing_in_ref.to_excel(w, index=False, sheet_name='غير موجود ببرنامج الرعاية')
        missing_in_site.to_excel(w, index=False, sheet_name='غير موجود بالموقع')
        amount_diff.to_excel(w, index=False, sheet_name='اختلاف مبلغ فقط')
        matches.to_excel(w, index=False, sheet_name='مطابق')

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
    }
    return summary, needs_update, missing_in_ref, amount_diff, missing_in_site


class App:
    def __init__(self, root):
        self.root = root
        self.root.title('مقارنة كفالات الايتام')
        self.root.state('zoomed')

        self.ref_var = tk.StringVar()
        self.site_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.needs_update_df = pd.DataFrame()
        self.missing_df = pd.DataFrame()
        self.amount_diff_df = pd.DataFrame()
        self.missing_site_df = pd.DataFrame()
        self.current_idx = 0
        self.review_df = pd.DataFrame()
        self.confirm_auto_var = tk.BooleanVar(value=True)
        self.enable_temp_save_var = tk.BooleanVar(value=False)
        self.allow_zero_update_var = tk.BooleanVar(value=False)

        self.build_ui()

    def build_ui(self):
        default_font = tkfont.nametofont('TkDefaultFont')
        default_font.configure(size=12, family='Tahoma')
        text_font = tkfont.nametofont('TkTextFont')
        text_font.configure(size=12, family='Tahoma')
        self.root.configure(bg='#f3f4f6')

        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass

        # Inspired by Karama system colors
        self.c_bg = '#e5e7eb'          # page background
        self.c_panel = '#f3f4f6'       # cards
        self.c_primary = '#2f4358'     # dark navy
        self.c_primary_active = '#24374a'
        self.c_secondary = '#5f7ea0'   # slate blue
        self.c_secondary_active = '#4d6b8d'
        self.c_beige = '#d9cfb1'
        self.c_beige_active = '#cdbf99'
        self.c_text_dark = '#111827'

        self.root.configure(bg=self.c_bg)

        self.main_frame = tk.Frame(self.root, bg=self.c_bg)

        style.configure('Treeview',
                        font=('Tahoma', 11),
                        rowheight=30,
                        background='#ffffff',
                        fieldbackground='#ffffff',
                        foreground=self.c_text_dark)
        style.configure('Treeview.Heading',
                        font=('Tahoma', 11, 'bold'),
                        background=self.c_primary,
                        foreground='white')
        style.map('Treeview', background=[('selected', '#1d4ed8')], foreground=[('selected', '#ffffff')])

        # Button palette
        self.btn_primary = self.c_primary
        self.btn_primary_active = self.c_primary_active
        self.btn_secondary = self.c_secondary
        self.btn_secondary_active = self.c_secondary_active
        self.btn_beige = self.c_beige
        self.btn_beige_active = self.c_beige_active

        pad = {'padx': 8, 'pady': 3}

        tk.Label(self.main_frame, text='ملف برنامج الرعاية (ايتام من البرنامج):', anchor='e').grid(row=0, column=2, sticky='e', **pad)
        tk.Entry(self.main_frame, textvariable=self.ref_var, width=70, justify='right').grid(row=0, column=1, **pad)
        tk.Button(self.main_frame, text='اختيار', command=self.pick_ref, bg=self.btn_secondary, fg='white', activebackground=self.btn_secondary_active, activeforeground='white').grid(row=0, column=0, **pad)

        tk.Label(self.main_frame, text='ملف الموقع المصدّر:', anchor='e').grid(row=1, column=2, sticky='e', **pad)
        tk.Entry(self.main_frame, textvariable=self.site_var, width=70, justify='right').grid(row=1, column=1, **pad)
        tk.Button(self.main_frame, text='اختيار', command=self.pick_site, bg=self.btn_secondary, fg='white', activebackground=self.btn_secondary_active, activeforeground='white').grid(row=1, column=0, **pad)

        tk.Label(self.main_frame, text='ملف النتيجة:', anchor='e').grid(row=2, column=2, sticky='e', **pad)
        tk.Entry(self.main_frame, textvariable=self.out_var, width=70, justify='right').grid(row=2, column=1, **pad)
        tk.Button(self.main_frame, text='حفظ باسم', command=self.pick_out, bg=self.btn_secondary, fg='white', activebackground=self.btn_secondary_active, activeforeground='white').grid(row=2, column=0, **pad)

        # صف أزرار التصفية والعرض
        filter_frame = tk.LabelFrame(self.main_frame, text='تصفية النتائج', font=('Tahoma', 10, 'bold'), fg=self.c_primary, bg=self.c_panel, padx=5, pady=5)
        filter_frame.grid(row=3, column=0, columnspan=3, padx=8, pady=6, sticky='ew')
        
        tk.Button(filter_frame, text='عرض كل الحالات', command=lambda: self.set_review_mode('all'), width=16, bg=self.btn_secondary, fg='white', activebackground=self.btn_secondary_active, activeforeground='white').grid(row=0, column=0, padx=3, pady=2)
        tk.Button(filter_frame, text='غير موجود ببرنامج الرعاية', command=lambda: self.set_review_mode('missing'), width=22, bg=self.btn_secondary, fg='white', activebackground=self.btn_secondary_active, activeforeground='white').grid(row=0, column=1, padx=3, pady=2)
        tk.Button(filter_frame, text='اختلاف مبلغ فقط', command=lambda: self.set_review_mode('amount_diff'), width=16, bg=self.btn_secondary, fg='white', activebackground=self.btn_secondary_active, activeforeground='white').grid(row=0, column=2, padx=3, pady=2)
        tk.Button(filter_frame, text='غير موجود بالموقع', command=lambda: self.set_review_mode('missing_site'), width=18, bg=self.btn_secondary, fg='white', activebackground=self.btn_secondary_active, activeforeground='white').grid(row=0, column=3, padx=3, pady=2)
        
        # صف الأزرار الرئيسية
        action_frame = tk.Frame(self.main_frame, bg=self.c_bg)
        action_frame.grid(row=4, column=0, columnspan=3, padx=8, pady=8, sticky='ew')
        
        tk.Button(action_frame, text='ابدأ المقارنة', command=self.run_compare, width=18, height=2, bg=self.btn_primary, fg='white', activebackground=self.btn_primary_active, activeforeground='white', font=('Tahoma', 11, 'bold')).pack(side='left', padx=4)
        tk.Button(action_frame, text='تنفيذ التعديل التلقائي', command=self.run_auto_update, width=22, height=2, bg=self.c_beige, fg='#333', activebackground=self.btn_beige_active, activeforeground='#333', font=('Tahoma', 11, 'bold')).pack(side='left', padx=4)
        tk.Button(action_frame, text='خروج', command=self.root.destroy, width=12, height=2, bg='#dc2626', fg='white', activebackground='#b91c1c', activeforeground='white', font=('Tahoma', 11, 'bold')).pack(side='right', padx=4)

        self.status = tk.Label(self.main_frame, text='جاهز', fg='darkgreen', anchor='w', justify='left')
        self.status.grid(row=5, column=1, sticky='w', padx=8, pady=2)

        # Inline review panel (same window)
        panel = tk.LabelFrame(self.main_frame, text='مراجعة اسم باسم', padx=8, pady=8, bg=self.c_panel)
        panel.grid(row=6, column=0, columnspan=3, padx=8, pady=6, sticky='ew')

        self.lbl_counter = tk.Label(panel, text='لا توجد بيانات بعد', bg=self.c_panel, anchor='e', justify='right', fg=self.c_primary)
        self.lbl_counter.grid(row=0, column=0, columnspan=4, sticky='e', pady=4)

        self.v_nat = tk.StringVar()
        self.v_name_site = tk.StringVar()
        self.v_name_ref = tk.StringVar()
        self.v_amount_site = tk.StringVar()
        self.v_amount_expected = tk.StringVar()
        self.v_reason = tk.StringVar()

        rows = [
            ('الرقم الوطني', self.v_nat),
            ('الاسم في الموقع', self.v_name_site),
            ('الاسم في برنامج الرعاية', self.v_name_ref),
            ('المبلغ الحالي', self.v_amount_site),
            ('المبلغ الفعلي', self.v_amount_expected),
            ('السبب', self.v_reason),
        ]
        for i, (title, var) in enumerate(rows):
            # RTL: label on the right side, field to the left
            tk.Label(panel, text=title, width=22, anchor='e', bg=self.c_panel, fg=self.c_text_dark).grid(row=i + 1, column=3, sticky='e', pady=3, padx=(4, 2))
            tk.Entry(panel, textvariable=var, width=82, justify='right').grid(row=i + 1, column=0, columnspan=3, sticky='ew', padx=(2, 8), pady=3)

        # RTL buttons order (right -> left)
        tk.Button(panel, text='تم التعديل - التالي', command=lambda: self.move_review(1), width=18, bg=self.btn_primary, fg='white', activebackground=self.btn_primary_active, activeforeground='white').grid(row=8, column=3, pady=8, sticky='e')
        tk.Button(panel, text='السابق', command=lambda: self.move_review(-1), width=12, bg=self.btn_secondary, fg='white', activebackground=self.btn_secondary_active, activeforeground='white').grid(row=8, column=2, pady=8, sticky='w')
        tk.Button(panel, text='نسخ الرقم الوطني', command=self.copy_nat, width=18, bg=self.btn_secondary, fg='white', activebackground=self.btn_secondary_active, activeforeground='white').grid(row=8, column=0, pady=8, sticky='w')
        tk.Button(panel, text='خروج', command=self.root.destroy, width=12, bg=self.btn_secondary, fg='white', activebackground=self.btn_secondary_active, activeforeground='white').grid(row=8, column=1, pady=8)

        panel.grid_columnconfigure(0, weight=1)
        panel.grid_columnconfigure(1, weight=1)
        panel.grid_columnconfigure(2, weight=1)

        list_frame = tk.LabelFrame(self.main_frame, text='قائمة الحالات في وضع العرض الحالي', padx=6, pady=6, bg=self.c_panel)
        list_frame.grid(row=7, column=0, columnspan=3, padx=8, pady=6, sticky='nsew')
        self.main_frame.grid_rowconfigure(7, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=1)

        # RTL-friendly order: reason at right, nat at left
        cols = ('reason', 'amount_expected', 'amount_site', 'name_site', 'nat')
        self.tree = ttk.Treeview(list_frame, columns=cols, show='headings', height=10)
        self.tree.heading('reason', text='السبب')
        self.tree.heading('amount_expected', text='المبلغ الفعلي')
        self.tree.heading('amount_site', text='المبلغ الحالي')
        self.tree.heading('name_site', text='الاسم في الموقع')
        self.tree.heading('nat', text='الرقم الوطني')

        self.tree.column('reason', width=220, anchor='e')
        self.tree.column('amount_expected', width=130, anchor='e')
        self.tree.column('amount_site', width=130, anchor='e')
        self.tree.column('name_site', width=380, anchor='e')
        self.tree.column('nat', width=170, anchor='e')

        ys = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=ys.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        ys.grid(row=0, column=1, sticky='ns')
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)

        self.main_frame.place(relx=0, rely=0, relwidth=1, relheight=1)

    def pick_ref(self):
        p = filedialog.askopenfilename(filetypes=[('Excel files', '*.xls *.xlsx')])
        if p:
            self.ref_var.set(p)
            self.suggest_output()

    def pick_site(self):
        p = filedialog.askopenfilename(filetypes=[('Excel files', '*.xls *.xlsx')])
        if p:
            self.site_var.set(p)
            self.suggest_output()

    def pick_out(self):
        p = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')])
        if p:
            self.out_var.set(p)

    def suggest_output(self):
        site = self.site_var.get().strip()
        if site and not self.out_var.get().strip():
            base = os.path.splitext(site)[0]
            ts = datetime.now().strftime('%Y%m%d_%H%M')
            self.out_var.set(f"{base}_compare_result_{ts}.xlsx")

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

            s, needs_update, missing_df, amount_diff_df, missing_site_df = compare_and_export(ref, site, out)
            self.needs_update_df = needs_update.reset_index(drop=True)
            self.missing_df = missing_df.reset_index(drop=True)
            self.amount_diff_df = amount_diff_df.reset_index(drop=True)
            self.missing_site_df = missing_site_df.reset_index(drop=True)
            self.current_idx = 0
            self.set_review_mode('all')

            msg = (
                f"تمت المقارنة بنجاح\n\n"
                f"سجلات الموقع: {s['site_rows']}\n"
                f"سجلات برنامج الرعاية: {s['ref_rows']}\n"
                f"يحتاج تعديل: {s['needs_update']}\n"
                f"مطابق: {s['matches']}\n"
                f"غير موجود ببرنامج الرعاية (تعديل الى 0): {s['missing_in_ref_updates_to_zero']}\n"
                f"موجود ببرنامج الرعاية وغير موجود بالموقع (بحاجة اضافة): {s['missing_in_site_needs_add']}\n"
                f"اختلاف مبلغ فقط: {s['amount_diff_only']}\n"
                f"تكرار ببرنامج الرعاية: {s['ref_duplicates']}\n"
                f"تكرار بالموقع: {s['site_duplicates']}\n\n"
                f"ملف النتيجة:\n{out}"
            )
            self.status.config(text='تمت العملية بنجاح')
            messagebox.showinfo('نجاح', msg)
        except Exception as e:
            self.status.config(text='حدث خطأ')
            messagebox.showerror('خطأ', str(e))

    def set_review_mode(self, mode='all'):
        if mode == 'missing':
            self.review_df = self.missing_df.copy()
            mode_title = 'غير موجود ببرنامج الرعاية (تصفير)'
        elif mode == 'missing_site':
            self.review_df = self.missing_site_df.copy()
            mode_title = 'موجود ببرنامج الرعاية وغير موجود بالموقع'
        elif mode == 'amount_diff':
            self.review_df = self.amount_diff_df.copy()
            mode_title = 'اختلاف مبلغ فقط'
        else:
            self.review_df = self.needs_update_df.copy()
            mode_title = 'كل الحالات'
        self.current_idx = 0
        self.status.config(text=f'وضع العرض: {mode_title}')
        self.populate_tree()
        self.fill_review()

    def populate_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if self.review_df.empty:
            return
        for _, row in self.review_df.iterrows():
            reason = str(row.get('سبب', ''))
            tag = 'amount_diff'
            if 'غير موجود ببرنامج الرعاية' in reason:
                tag = 'missing'
            elif 'غير موجود بالموقع' in reason:
                tag = 'missing_site'
            self.tree.insert('', 'end', values=(
                reason,
                str(row.get('المبلغ_الفعلي', '')),
                str(row.get('المبلغ_في_الموقع', '')),
                str(row.get('الاسم_في_الموقع', '')),
                str(row.get('الرقم_الوطني', '')),
            ), tags=(tag,))
        self.tree.tag_configure('missing', background='#fef3c7')   # amber
        self.tree.tag_configure('amount_diff', background='#dbeafe')  # blue
        self.tree.tag_configure('missing_site', background='#dcfce7')  # green

    def fill_review(self):
        if self.review_df.empty:
            self.lbl_counter.config(text='لا توجد بيانات لهذا النوع بعد')
            self.v_nat.set('')
            self.v_name_site.set('')
            self.v_name_ref.set('')
            self.v_amount_site.set('')
            self.v_amount_expected.set('')
            self.v_reason.set('')
            return
        n = len(self.review_df)
        i = self.current_idx
        row = self.review_df.iloc[i]
        self.lbl_counter.config(text=f'الحالة {i + 1} من {n}')
        self.v_nat.set(str(row.get('الرقم_الوطني', '')))
        self.v_name_site.set(str(row.get('الاسم_في_الموقع', '')))
        self.v_name_ref.set(str(row.get('الاسم_في_برنامج_الرعاية', '')))
        self.v_amount_site.set(str(row.get('المبلغ_في_الموقع', '')))
        self.v_amount_expected.set(str(row.get('المبلغ_الفعلي', '')))
        self.v_reason.set(str(row.get('سبب', '')))
        items = self.tree.get_children()
        if items and i < len(items):
            self.tree.selection_set(items[i])
            self.tree.focus(items[i])
            self.tree.see(items[i])

    def move_review(self, delta):
        if self.review_df.empty:
            return
        new_idx = self.current_idx + delta
        if new_idx < 0:
            new_idx = 0
        if new_idx >= len(self.review_df):
            messagebox.showinfo('تم', 'انتهت جميع الحالات.')
            new_idx = len(self.review_df) - 1
        self.current_idx = new_idx
        self.fill_review()

    def copy_nat(self):
        nat = self.v_nat.get().strip()
        if not nat:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(nat)

    def run_auto_update(self):
        out_file = self.out_var.get().strip()
        if not out_file:
            messagebox.showerror('خطأ', 'يرجى تنفيذ المقارنة أولاً أو اختيار ملف النتيجة.')
            return
        if not os.path.exists(out_file):
            messagebox.showerror('خطأ', f'ملف النتيجة غير موجود:\n{out_file}')
            return

        try:
            import importlib.util
            spec = importlib.util.spec_from_file_location("auto_update_gui", 
                                                          resource_path('auto_update_gui.py'))
            gui_module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(gui_module)
            
            # إخفاء الواجهة الرئيسية
            self.main_frame.place_forget()
            
            # إنشاء إطار التعديلات في نفس النافذة
            self.update_frame = tk.Frame(self.root, bg=self.c_bg)
            self.update_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
            
            options = {
                'excel_path': out_file,
                'on_back': self._show_main_frame,  # دالة العودة للرئيسية
                'on_exit': self.root.destroy,       # دالة الخروج
            }
            
            self.update_app = gui_module.AutoUpdateGUI(self.update_frame, out_file, options)
        except Exception as e:
            self.main_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
            messagebox.showerror('خطأ', f'تعذر فتح نافذة التحديث:\n{e}')

    def _show_main_frame(self):
        """العودة للصفحة الرئيسية"""
        if hasattr(self, 'update_frame'):
            self.update_frame.destroy()
        self.main_frame.place(relx=0, rely=0, relwidth=1, relheight=1)

    def on_tree_select(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        idx = self.tree.index(item)
        if idx != self.current_idx and 0 <= idx < len(self.review_df):
            self.current_idx = idx
            self.fill_review()


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
