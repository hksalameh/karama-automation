import base64
import ctypes
import json
import os
import sys
import subprocess
import threading
import tkinter as tk
from ctypes import wintypes
from tkinter import ttk, messagebox, scrolledtext
from pathlib import Path
from datetime import datetime
from shutil import which


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


def writable_app_dir():
    base_dir = os.environ.get('LOCALAPPDATA') or os.environ.get('APPDATA') or os.path.expanduser('~')
    path = os.path.join(base_dir, 'KafalaCompareApp')
    os.makedirs(path, exist_ok=True)
    return path


class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ('cbData', wintypes.DWORD),
        ('pbData', ctypes.POINTER(ctypes.c_byte)),
    ]


def _dpapi_protect(text):
    """Encrypt text for the current Windows user using DPAPI."""
    if os.name != 'nt':
        raise RuntimeError('حفظ كلمة المرور الآمن متاح على Windows فقط.')

    raw = text.encode('utf-8')
    in_buffer = ctypes.create_string_buffer(raw)
    in_blob = DATA_BLOB(len(raw), ctypes.cast(in_buffer, ctypes.POINTER(ctypes.c_byte)))
    out_blob = DATA_BLOB()

    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    if not crypt32.CryptProtectData(
        ctypes.byref(in_blob), None, None, None, None, 0, ctypes.byref(out_blob)
    ):
        raise ctypes.WinError()

    try:
        encrypted = ctypes.string_at(out_blob.pbData, out_blob.cbData)
        return base64.b64encode(encrypted).decode('ascii')
    finally:
        kernel32.LocalFree(out_blob.pbData)


def _dpapi_unprotect(encoded):
    """Decrypt text previously protected for the current Windows user."""
    if os.name != 'nt':
        return ''

    encrypted = base64.b64decode(encoded.encode('ascii'))
    in_buffer = ctypes.create_string_buffer(encrypted)
    in_blob = DATA_BLOB(len(encrypted), ctypes.cast(in_buffer, ctypes.POINTER(ctypes.c_byte)))
    out_blob = DATA_BLOB()

    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    if not crypt32.CryptUnprotectData(
        ctypes.byref(in_blob), None, None, None, None, 0, ctypes.byref(out_blob)
    ):
        raise ctypes.WinError()

    try:
        raw = ctypes.string_at(out_blob.pbData, out_blob.cbData)
        return raw.decode('utf-8')
    finally:
        kernel32.LocalFree(out_blob.pbData)


class AutoUpdateGUI:
    """واجهة الدخول إلى كرامة والتعديل ثم الحفظ المؤقت فقط."""

    def __init__(self, parent_frame, excel_path, options):
        self.root = parent_frame
        self.excel_path = excel_path
        self.options = options or {}
        self.runtime_dir = writable_app_dir()
        self.credentials_path = os.path.join(self.runtime_dir, 'karama_credentials.json')
        self.process = None
        self.reader_thread = None
        self.running = False
        self.completed_temp_save = False
        self.stop_requested = False

        now = datetime.now()
        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.show_password_var = tk.BooleanVar(value=False)
        self.remember_credentials_var = tk.BooleanVar(value=True)
        self.year_var = tk.StringVar(value=str(now.year))
        self.month_var = tk.StringVar(value=str(now.month))
        self.category_var = tk.StringVar(value='ايتام')

        self.progress_var = tk.DoubleVar(value=0)
        self.current_nat_var = tk.StringVar(value='---')
        self.current_change_var = tk.StringVar(value='---')
        self.current_status_var = tk.StringVar(value='جاهز')
        self.total_var = tk.StringVar(value='---')

        self._load_saved_credentials()
        self._build_ui()
        self._validate_excel_exists()

    def _load_saved_credentials(self):
        if not os.path.isfile(self.credentials_path):
            return
        try:
            with open(self.credentials_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            username = str(data.get('username', '')).strip()
            encrypted = str(data.get('password_dpapi', '')).strip()
            password = _dpapi_unprotect(encrypted) if encrypted else ''
            if username and password:
                self.username_var.set(username)
                self.password_var.set(password)
                self.remember_credentials_var.set(True)
        except Exception:
            # لا نمنع تشغيل البرنامج إذا تلف ملف الإعدادات أو تغير مستخدم Windows.
            self.username_var.set('')
            self.password_var.set('')

    def _save_credentials(self, username, password):
        if not self.remember_credentials_var.get():
            self._delete_saved_credentials()
            return
        data = {
            'version': 1,
            'username': username,
            'password_dpapi': _dpapi_protect(password),
        }
        temp_path = self.credentials_path + '.tmp'
        with open(temp_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(temp_path, self.credentials_path)

    def _delete_saved_credentials(self):
        try:
            if os.path.exists(self.credentials_path):
                os.remove(self.credentials_path)
        except Exception:
            pass

    def _clear_saved_credentials(self):
        if self.running:
            return
        self._delete_saved_credentials()
        self.username_var.set('')
        self.password_var.set('')
        self.remember_credentials_var.set(True)
        messagebox.showinfo('تم', 'تم مسح بيانات دخول كرامة المحفوظة من هذا الجهاز.')

    def _build_ui(self):
        self.root.configure(bg='#e5e7eb')

        header = tk.Frame(self.root, bg='#2f4358', height=82)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(
            header,
            text='التعديل التلقائي في موقع كرامة',
            font=('Tahoma', 16, 'bold'),
            fg='white', bg='#2f4358'
        ).pack(pady=(12, 3))
        tk.Label(
            header,
            text='التعديل الآلي ينتهي عند الحفظ المؤقت فقط — الحفظ النهائي يبقى يدوياً',
            font=('Tahoma', 10), fg='#dbeafe', bg='#2f4358'
        ).pack()

        content = tk.Frame(self.root, bg='#e5e7eb')
        content.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        settings = tk.LabelFrame(
            content, text='إعدادات التشغيل',
            font=('Tahoma', 11, 'bold'), fg='#2f4358', bg='white', padx=10, pady=8
        )
        settings.pack(fill=tk.X, pady=(0, 8))

        tk.Label(settings, text='ملف المقارنة:', bg='white', font=('Tahoma', 10, 'bold')).grid(row=0, column=5, sticky='e', padx=5, pady=4)
        self.file_label = tk.Label(settings, text=Path(self.excel_path).name, bg='white', fg='#334155', anchor='e', font=('Tahoma', 10))
        self.file_label.grid(row=0, column=0, columnspan=5, sticky='ew', padx=5, pady=4)

        tk.Label(settings, text='اسم المستخدم:', bg='white', font=('Tahoma', 10, 'bold')).grid(row=1, column=5, sticky='e', padx=5, pady=4)
        self.username_entry = tk.Entry(settings, textvariable=self.username_var, justify='right', width=24, font=('Tahoma', 10))
        self.username_entry.grid(row=1, column=4, sticky='ew', padx=5, pady=4)

        tk.Label(settings, text='كلمة المرور:', bg='white', font=('Tahoma', 10, 'bold')).grid(row=1, column=3, sticky='e', padx=5, pady=4)
        self.password_entry = tk.Entry(settings, textvariable=self.password_var, justify='right', show='●', width=24, font=('Tahoma', 10))
        self.password_entry.grid(row=1, column=2, sticky='ew', padx=5, pady=4)
        self.show_password_check = tk.Checkbutton(
            settings, text='إظهار', variable=self.show_password_var,
            command=self._toggle_password, bg='white', font=('Tahoma', 9)
        )
        self.show_password_check.grid(row=1, column=1, sticky='w', padx=5)

        tk.Label(settings, text='السنة:', bg='white', font=('Tahoma', 10, 'bold')).grid(row=2, column=5, sticky='e', padx=5, pady=4)
        self.year_entry = tk.Entry(settings, textvariable=self.year_var, justify='center', width=10, font=('Tahoma', 10))
        self.year_entry.grid(row=2, column=4, sticky='w', padx=5, pady=4)

        tk.Label(settings, text='الشهر:', bg='white', font=('Tahoma', 10, 'bold')).grid(row=2, column=3, sticky='e', padx=5, pady=4)
        self.month_combo = ttk.Combobox(settings, textvariable=self.month_var, values=[str(i) for i in range(1, 13)], width=8, state='readonly', justify='center')
        self.month_combo.grid(row=2, column=2, sticky='w', padx=5, pady=4)

        tk.Label(settings, text='التصنيف:', bg='white', font=('Tahoma', 10, 'bold')).grid(row=2, column=1, sticky='e', padx=5, pady=4)
        self.category_combo = ttk.Combobox(
            settings, textvariable=self.category_var,
            values=['ايتام', 'اسر', 'طلاب علم'], width=14,
            state='readonly', justify='center'
        )
        self.category_combo.grid(row=2, column=0, sticky='w', padx=5, pady=4)

        credential_row = tk.Frame(settings, bg='white')
        credential_row.grid(row=3, column=0, columnspan=6, sticky='ew', padx=5, pady=(5, 0))
        self.remember_check = tk.Checkbutton(
            credential_row,
            text='حفظ بيانات الدخول على هذا الجهاز',
            variable=self.remember_credentials_var,
            bg='white', font=('Tahoma', 9)
        )
        self.remember_check.pack(side=tk.RIGHT)
        self.clear_credentials_btn = tk.Button(
            credential_row, text='مسح المحفوظ', command=self._clear_saved_credentials,
            font=('Tahoma', 8), padx=8, pady=1
        )
        self.clear_credentials_btn.pack(side=tk.RIGHT, padx=8)
        tk.Label(
            credential_row,
            text='كلمة المرور تُحفظ مشفّرة بحماية Windows ولا تُخزن كنص عادي.',
            bg='white', fg='#64748b', font=('Tahoma', 9)
        ).pack(side=tk.RIGHT, padx=8)

        for col in range(6):
            settings.grid_columnconfigure(col, weight=1 if col in (0, 2, 4) else 0)

        actions = tk.Frame(content, bg='#e5e7eb')
        actions.pack(fill=tk.X, pady=5)

        self.start_btn = tk.Button(
            actions, text='بدء التعديل التلقائي', command=self.start_automation,
            bg='#2f4358', fg='white', activebackground='#24374a', activeforeground='white',
            font=('Tahoma', 11, 'bold'), padx=20, pady=9
        )
        self.start_btn.pack(side=tk.LEFT, padx=4)

        self.cancel_btn = tk.Button(
            actions, text='إلغاء التشغيل', command=self.cancel_automation,
            bg='#dc2626', fg='white', activebackground='#b91c1c', activeforeground='white',
            font=('Tahoma', 10, 'bold'), padx=16, pady=9, state=tk.DISABLED
        )
        self.cancel_btn.pack(side=tk.LEFT, padx=4)

        self.back_btn = tk.Button(
            actions, text='العودة للرئيسية', command=self.go_back,
            bg='#5f7ea0', fg='white', activebackground='#4d6b8d', activeforeground='white',
            font=('Tahoma', 10, 'bold'), padx=16, pady=9
        )
        self.back_btn.pack(side=tk.RIGHT, padx=4)

        progress_box = tk.LabelFrame(
            content, text='حالة التنفيذ',
            font=('Tahoma', 11, 'bold'), fg='#2f4358', bg='white', padx=10, pady=8
        )
        progress_box.pack(fill=tk.X, pady=8)

        self.progress_bar = ttk.Progressbar(progress_box, variable=self.progress_var, maximum=100, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(2, 8))

        row = tk.Frame(progress_box, bg='white')
        row.pack(fill=tk.X)
        tk.Label(row, text='الرقم الوطني:', bg='white', fg='#2f4358', font=('Tahoma', 10, 'bold')).pack(side=tk.RIGHT, padx=5)
        tk.Label(row, textvariable=self.current_nat_var, bg='white', fg='#111827', font=('Tahoma', 10)).pack(side=tk.RIGHT, padx=5)
        tk.Label(row, text='التعديل:', bg='white', fg='#2f4358', font=('Tahoma', 10, 'bold')).pack(side=tk.RIGHT, padx=(25, 5))
        tk.Label(row, textvariable=self.current_change_var, bg='white', fg='#111827', font=('Tahoma', 10)).pack(side=tk.RIGHT, padx=5)

        row2 = tk.Frame(progress_box, bg='white')
        row2.pack(fill=tk.X, pady=(6, 0))
        tk.Label(row2, text='الحالة:', bg='white', fg='#2f4358', font=('Tahoma', 10, 'bold')).pack(side=tk.RIGHT, padx=5)
        self.status_label = tk.Label(row2, textvariable=self.current_status_var, bg='white', fg='#111827', font=('Tahoma', 10), anchor='e')
        self.status_label.pack(side=tk.RIGHT, padx=5)
        tk.Label(row2, text='المجموع بعد الحفظ المؤقت:', bg='white', fg='#2f4358', font=('Tahoma', 10, 'bold')).pack(side=tk.LEFT, padx=(5, 2))
        tk.Label(row2, textvariable=self.total_var, bg='white', fg='#16a34a', font=('Tahoma', 11, 'bold')).pack(side=tk.LEFT, padx=5)

        log_box = tk.LabelFrame(
            content, text='سجل العملية',
            font=('Tahoma', 11, 'bold'), fg='#2f4358', bg='white', padx=6, pady=6
        )
        log_box.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        self.log_text = scrolledtext.ScrolledText(
            log_box, height=12, wrap=tk.WORD,
            font=('Tahoma', 9), bg='#f8fafc', fg='#111827'
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)

    def _validate_excel_exists(self):
        if not self.excel_path or not os.path.exists(self.excel_path):
            self.current_status_var.set('ملف المقارنة غير موجود')
            self.start_btn.config(state=tk.DISABLED)
            self._log(f'ملف المقارنة غير موجود: {self.excel_path}', 'ERROR')
        else:
            self._log(f'جاهز. ملف المقارنة: {self.excel_path}', 'INFO')
            if self.username_var.get() and self.password_var.get():
                self._log('تم تحميل بيانات دخول كرامة المحفوظة على هذا الجهاز.', 'INFO')

    def _toggle_password(self):
        self.password_entry.config(show='' if self.show_password_var.get() else '●')

    def _log(self, message, level='INFO'):
        if threading.current_thread() is not threading.main_thread():
            self.root.after(0, lambda: self._log(message, level))
            return
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f'[{timestamp}] {message}\n')
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _find_node(self):
        if getattr(sys, 'frozen', False):
            bundled = resource_path('node.exe')
            if os.path.isfile(bundled):
                return bundled
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            bundled = os.path.join(base_dir, 'KafalaCompareApp_build', 'node.exe')
            if os.path.isfile(bundled):
                return bundled
        return which('node')

    def _set_controls_running(self, running):
        self.running = running
        state = tk.DISABLED if running else tk.NORMAL
        self.start_btn.config(state=state)
        self.username_entry.config(state=state)
        self.password_entry.config(state=state)
        self.year_entry.config(state=state)
        self.show_password_check.config(state=state)
        self.remember_check.config(state=state)
        self.clear_credentials_btn.config(state=state)
        self.month_combo.config(state='disabled' if running else 'readonly')
        self.category_combo.config(state='disabled' if running else 'readonly')
        self.cancel_btn.config(state=tk.NORMAL if running else tk.DISABLED)

    def start_automation(self):
        if self.running:
            return

        username = self.username_var.get().strip()
        password = self.password_var.get()
        year = self.year_var.get().strip()
        month = self.month_var.get().strip()
        category = self.category_var.get().strip()

        if not username or not password:
            messagebox.showerror('بيانات ناقصة', 'أدخل اسم المستخدم وكلمة المرور لموقع كرامة.')
            return
        if len(year) != 4 or not year.isdigit():
            messagebox.showerror('بيانات غير صحيحة', 'أدخل السنة بأربعة أرقام.')
            return
        if not month.isdigit() or not 1 <= int(month) <= 12:
            messagebox.showerror('بيانات غير صحيحة', 'اختر الشهر الصحيح.')
            return
        if category not in ('ايتام', 'اسر', 'طلاب علم'):
            messagebox.showerror('بيانات غير صحيحة', 'اختر التصنيف الصحيح.')
            return

        try:
            self._save_credentials(username, password)
        except Exception as exc:
            if self.remember_credentials_var.get():
                messagebox.showerror('تعذر حفظ بيانات الدخول', f'لم أستطع حفظ بيانات الدخول بأمان:\n{exc}')
                return

        if not messagebox.askyesno(
            'تأكيد بدء التشغيل',
            f'سيتم العمل على:\nالسنة: {year}\nالشهر: {month}\nالتصنيف: {category}\n\n'
            'سيتم الضغط على "حفظ مؤقت" فقط بعد نجاح جميع التعديلات.\n'
            'لن يتم الضغط على "حفظ نهائي".\n\nهل تريد البدء؟'
        ):
            return

        node_path = self._find_node()
        script_path = resource_path('auto_update_from_diff.js')
        if not node_path:
            messagebox.showerror('خطأ', 'تعذر العثور على Node.js المطلوب لتشغيل الأتمتة.')
            return
        if not os.path.exists(script_path):
            messagebox.showerror('خطأ', f'سكربت الأتمتة غير موجود:\n{script_path}')
            return

        env = os.environ.copy()
        env['KARAMA_EXCEL_PATH'] = self.excel_path
        env['KARAMA_USERNAME'] = username
        env['KARAMA_PASSWORD'] = password
        env['KARAMA_YEAR'] = year
        env['KARAMA_MONTH'] = month
        env['KARAMA_CATEGORY'] = category
        env['KARAMA_OUTPUT_DIR'] = self.runtime_dir

        self.completed_temp_save = False
        self.stop_requested = False
        self.progress_var.set(0)
        self.total_var.set('---')
        self.current_nat_var.set('---')
        self.current_change_var.set('---')
        self.current_status_var.set('جاري تشغيل المتصفح...')
        self._set_controls_running(True)
        self._log(f'بدء التشغيل: {year}/{month} - {category}', 'INFO')

        creationflags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        try:
            self.process = subprocess.Popen(
                [node_path, script_path],
                cwd=self.runtime_dir,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                text=True,
                encoding='utf-8',
                errors='replace',
                creationflags=creationflags,
            )
        except Exception as exc:
            self._set_controls_running(False)
            messagebox.showerror('خطأ', f'تعذر تشغيل الأتمتة:\n{exc}')
            return

        self.reader_thread = threading.Thread(target=self._read_process_output, daemon=True)
        self.reader_thread.start()

    def _read_process_output(self):
        try:
            if not self.process or not self.process.stdout:
                return
            for raw_line in self.process.stdout:
                line = raw_line.strip()
                if not line:
                    continue
                if line.startswith('__PROGRESS__|'):
                    self._handle_progress_marker(line)
                elif line.startswith('__TEMP_SAVE_OK__|'):
                    self._handle_temp_save_ok(line)
                elif line.startswith('__FATAL__|'):
                    message = line.split('|', 1)[1] if '|' in line else line
                    self.root.after(0, lambda m=message: self._handle_fatal(m))
                else:
                    self._log(line, 'INFO')
            return_code = self.process.wait()
            self.root.after(0, lambda rc=return_code: self._process_finished(rc))
        except Exception as exc:
            self.root.after(0, lambda: self._handle_fatal(str(exc)))

    def _handle_progress_marker(self, line):
        parts = line.split('|', 6)
        if len(parts) < 7:
            return
        _, index, total, nat_id, old_amount, new_amount, status = parts
        try:
            idx = int(index)
            total_num = int(total)
            percent = (idx / total_num) * 100 if total_num else 0
        except Exception:
            percent = 0
        self.root.after(0, lambda: self._update_progress(percent, nat_id, old_amount, new_amount, status, index, total))

    def _update_progress(self, percent, nat_id, old_amount, new_amount, status, index, total):
        self.progress_var.set(percent)
        self.current_nat_var.set(nat_id or '---')
        self.current_change_var.set(f'{old_amount or "فارغ"} ← {new_amount or "فارغ"}')
        self.current_status_var.set(f'[{index}/{total}] {status}')

    def _handle_temp_save_ok(self, line):
        parts = line.split('|', 2)
        total = parts[1] if len(parts) > 1 else ''
        report_path = parts[2] if len(parts) > 2 else ''
        self.root.after(0, lambda: self._mark_temp_save_success(total, report_path))

    def _mark_temp_save_success(self, total, report_path):
        self.completed_temp_save = True
        self.progress_var.set(100)
        self.total_var.set(total or 'غير متاح')
        self.current_status_var.set('تم الحفظ المؤقت بنجاح — راجع المجموع ثم احفظ نهائياً يدوياً')
        self.cancel_btn.config(text='إغلاق جلسة المتصفح')
        self._log('✅ تم الحفظ المؤقت بنجاح.', 'SUCCESS')
        if total:
            self._log(f'المجموع الظاهر في الموقع بعد الحفظ المؤقت: {total}', 'INFO')
        if report_path:
            self._log(f'تقرير التشغيل: {report_path}', 'INFO')
        self._log('🛑 لم ولن يتم الضغط على زر حفظ نهائي.', 'WARNING')
        messagebox.showinfo(
            'تم الحفظ المؤقت',
            'انتهت التعديلات وتم الحفظ المؤقت بنجاح.\n\n'
            f'المجموع الظاهر: {total or "غير متاح"}\n\n'
            'اترك المتصفح مفتوحاً، راجع المجموع، ثم نفّذ الحفظ النهائي بنفسك إذا كان صحيحاً.'
        )

    def _handle_fatal(self, message):
        self.current_status_var.set('توقفت العملية بسبب خطأ')
        self._log(f'❌ {message}', 'ERROR')
        if not self.completed_temp_save:
            messagebox.showerror(
                'تم إيقاف العملية للحماية',
                f'{message}\n\nلم يتم تنفيذ الحفظ النهائي بواسطة البرنامج.'
            )

    def _process_finished(self, return_code):
        self._set_controls_running(False)
        self.cancel_btn.config(text='إلغاء التشغيل')
        if self.stop_requested:
            self.current_status_var.set('تم إلغاء التشغيل')
            return
        if self.completed_temp_save:
            self.current_status_var.set('انتهت الجلسة. الحفظ النهائي بقي يدوياً.')
            return
        if return_code == 0:
            self.current_status_var.set('انتهت العملية')
        else:
            self.current_status_var.set(f'انتهت العملية برمز خطأ {return_code}')

    def cancel_automation(self):
        if not self.process or self.process.poll() is not None:
            return

        if self.completed_temp_save:
            if not messagebox.askyesno('إغلاق المتصفح', 'هل تريد إغلاق جلسة المتصفح الآن؟'):
                return
        else:
            if not messagebox.askyesno(
                'تأكيد الإلغاء',
                'هل تريد إيقاف العملية الآن؟\nأي تعديلات لم يتم حفظها مؤقتاً ستُترك بدون حفظ.'
            ):
                return

        self.stop_requested = True
        try:
            self.process.terminate()
        except Exception:
            pass
        self._log('تم طلب إيقاف جلسة الأتمتة.', 'WARNING')

    def go_back(self):
        if self.running and not self.completed_temp_save:
            messagebox.showwarning('العملية تعمل', 'ألغِ التشغيل أولاً قبل العودة للرئيسية.')
            return

        on_back = self.options.get('on_back')
        if self.completed_temp_save and self.process and self.process.poll() is None:
            self._log('العودة للرئيسية مع إبقاء المتصفح مفتوحاً للمراجعة.', 'INFO')
        if on_back:
            on_back()
        else:
            self.root.destroy()
