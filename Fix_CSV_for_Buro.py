#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Нормализация данных *.csv с разных рабочих мест «Бастион» для корректного использования
в комплексе с программном обеспечением «Личные пропуска».

Функционал:
- объединяет CSV из выбранной папки в один Excel,
- проверяет FULLCARDCODE (12 HEX),
- удаляет заглушки в виде русских наименований полей и пустые строки,
- удаляет полные дубликаты строк (одну из них),
- поля расставляются в правильном порядке, если отсутсвуют - создаются новые(пустые),
- Удаление начальных и конечных пробелов из всех строковых полей,
- проверяет FULLCARDCODE на дубликаты и удаляет все кроме первого с сохранением в отдельную таблицу списка дубликатов,
- переносит WORG6 → WORG7 если пусты WORG7 и WORG8, эти поля содержат название организаций,
- заполняет поле подразделение – WDEP8 ("Нет данных") если строка пустая,
- пересохраняет готовый файл через Excel COM для 100% совместимости с импортёром.
- добавлена проверка готового xlsx файла на соответствие структуры

Автор: Шаулис Э.Ю.
Дата: 01.03.2026
Версия: 1.5
"""

import os
import glob
import pandas as pd
import re
from datetime import datetime
from tkinter import Tk, Label, Button, Text, END, DISABLED, NORMAL, messagebox, filedialog, Menu, ttk, Scrollbar, Frame
from tkinter.font import Font
import threading

try:
    import win32com.client as win32
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

TARGET_FIELDS = [
    'B_VERSION', 'NAME', 'FIRSTNAME', 'SECONDNAME', 'TABLENO', 'FULLCARDCODE', 'ALNAME',
    'WDEP2', 'WDEP3', 'WDEP4', 'WDEP5', 'WDEP6', 'WDEP7', 'WDEP8',
    'WORG1', 'WORG2', 'WORG3', 'WORG4', 'WORG5', 'WORG6', 'WORG7', 'WORG8',
    'POST', 'PHONE', 'DOCTYPE', 'DOCSER', 'DOCNO', 'DOCISSUEORGAN', 'ADDRESS',
    'BIRTHPLACE', 'PERSONCAT', 'SITIZENSHIP', 'COMMENTS',
    'ADDFLD1', 'ADDFLD2', 'ADDFLD3', 'ADDFLD4', 'ADDFLD5', 'ADDFLD6', 'ADDFLD7',
    'ADDFLD8', 'ADDFLD9', 'ADDFLD10', 'ADDFLD11', 'ADDFLD12', 'ADDFLD13', 'ADDFLD14',
    'ADDFLD15', 'ADDFLD16', 'ADDFLD17', 'ADDFLD18', 'ADDFLD19', 'ADDFLD20',
    'SERIALNUMBER', 'MIFARE_SERIALNO', 'CORP_CODE', 'PASSKIND', 'PINCODE', 'PS_COMMENT',
    'PASSCC', 'RETURNREASON', 'PASSFORM', 'VISITGOAL', 'ACCEPTDEP', 'ACCEPTPERSON',
    'BLOCKEDREASON', 'PRIORITY', 'CARD_IDENTIFIER_TYPE_ID', 'CARDACC_TYPE_CARD',
    'CARDACC_ISSUE', 'PRIOR_MIFER_CARD_STATUS', 'PASSTYPE', 'IS_BLOCKED', 'SEX',
    'IS_PERSON_AGREEMENT_EXISTS', 'STARTDATE', 'ENDDATE', 'STARTTIME', 'ENDTIME',
    'DOCISSUEDATE', 'BIRTHDATE', 'ALTERDATE', 'CREATEDATE', 'RETURNDATE', 'PASSCDATE',
    'BLOCKEDDATA', 'PERSON_AGREEMENT_DATE', 'EMAIL'
]

class App:
    # Цветовая схема
    COLORS = {
        'bg': '#f0f2f5',
        'card_bg': '#ffffff',
        'primary': '#1890ff',
        'primary_hover': '#40a9ff',
        'success': '#52c41a',
        'warning': '#faad14',
        'error': '#ff4d4f',
        'text': '#303133',
        'text_secondary': '#606266',
        'border': '#d9d9d9',
    }

    def __init__(self, root):
        self.root = root
        root.title("📋 Нормализация CSV → Excel")
        root.geometry("900x750")
        root.resizable(True, True)
        root.configure(bg=self.COLORS['bg'])

        # Настройка шрифтов
        self.font_title = Font(family="Segoe UI", size=14, weight="bold")
        self.font_normal = Font(family="Segoe UI", size=10)
        self.font_log = Font(family="Consolas", size=9)

        self._create_menu()
        self._create_header()
        self._create_main_card()
        self._create_buttons()
        self._create_status_bar()

        # Центрирование окна на экране (выполняется после создания всех виджетов)
        self.root.update_idletasks()
        window_width = 900
        window_height = 750
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _ui(self, func, *args, wait=False, **kwargs):
        """Run UI operations safely from worker threads."""
        if threading.current_thread() is threading.main_thread():
            return func(*args, **kwargs)

        if wait:
            done = threading.Event()
            result = {}

            def wrapper():
                try:
                    result['value'] = func(*args, **kwargs)
                except Exception as e:
                    result['error'] = e
                finally:
                    done.set()

            self.root.after(0, wrapper)
            done.wait()
            if 'error' in result:
                raise result['error']
            return result.get('value')

        self.root.after(0, lambda: func(*args, **kwargs))

    def _create_menu(self):
        menu_bar = Menu(self.root)
        file_menu = Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="📁 Выбрать папку...", command=self.run_process)
        file_menu.add_command(label="✅ Проверить файл...", command=self.check_export_file)
        file_menu.add_separator()
        file_menu.add_command(label="❌ Выход", command=self.root.quit)
        menu_bar.add_cascade(label="Файл", menu=file_menu)
        
        help_menu = Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="ℹ️ О программе", command=self._show_about)
        menu_bar.add_cascade(label="Справка", menu=help_menu)
        
        self.root.config(menu=menu_bar)

    def _create_header(self):
        header = Frame(self.root, bg=self.COLORS['primary'], height=60)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        Label(header, text="🏢 Нормализация данных экспорта *.csv серверов «Бастион»",
              font=Font(family="Segoe UI", size=13, weight="bold"),
              bg=self.COLORS['primary'], fg="white").pack(pady=15)

    def _create_main_card(self):
        # Карточка с логом
        card = Frame(self.root, bg=self.COLORS['card_bg'], bd=0)
        card.pack(padx=20, pady=15, fill="both", expand=True)
        
        # Рамка с тенью (через borderwidth и relief)
        card.config(highlightbackground=self.COLORS['border'], highlightthickness=1)
        
        Label(card, text="📝 Журнал выполнения", font=self.font_title,
              bg=self.COLORS['card_bg'], fg=self.COLORS['text']).pack(anchor="w", padx=15, pady=(15, 5))

        # Контейнер для текста с прокруткой
        text_frame = Frame(card, bg=self.COLORS['card_bg'])
        text_frame.pack(padx=15, pady=5, fill="both", expand=True)

        self.log_text = Text(text_frame, wrap="word", font=self.font_log,
                             bg='#1e1e1e', fg='#d4d4d4', insertbackground='white',
                             relief="flat", height=18, padx=10, pady=10)
        self.log_text.pack(side="left", fill="both", expand=True)
        
        scrollbar = Scrollbar(text_frame, command=self.log_text.yview, bg=self.COLORS['border'])
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        self.log_text.config(state=DISABLED)

        # Прогресс-бар
        self.progress = ttk.Progressbar(card, mode='indeterminate', length=300)
        self.progress.pack(pady=(0, 10))

    def _create_buttons(self):
        btn_frame = Frame(self.root, bg=self.COLORS['bg'])
        btn_frame.pack(pady=(0, 15))

        self.btn_process = Button(btn_frame, text="🚀 Выбрать папку и обработать",
                                   font=self.font_normal, bg=self.COLORS['primary'], fg="white",
                                   activebackground=self.COLORS['primary_hover'], activeforeground="white",
                                   bd=0, padx=25, pady=10, cursor="hand2", command=self.run_process)
        self.btn_process.pack(side="left", padx=10)

        self.btn_check = Button(btn_frame, text="✅ Проверить готовый файл",
                                 font=self.font_normal, bg=self.COLORS['card_bg'], fg=self.COLORS['primary'],
                                 activebackground='#e6f7ff', activeforeground=self.COLORS['primary'],
                                 bd=1, relief="solid", padx=25, pady=10, cursor="hand2", command=self.check_export_file)
        self.btn_check.pack(side="left", padx=10)

    def _create_status_bar(self):
        self.status_var = self.root.var = "Готов к работе"
        self.status_label = Label(self.root, text=self.status_var, 
                                   font=self.font_normal, bg=self.COLORS['card_bg'],
                                   fg=self.COLORS['text_secondary'], anchor="w", padx=15, pady=8)
        self.status_label.pack(side="bottom", fill="x")

    def _show_about(self):
        messagebox.showinfo("О программе",
            "📋 Нормализация CSV → Excel\n\n"
            "Версия: 1.5\n"
            "Автор: Шаулис Э.Ю.\n"
            "Дата: 01.03.2026\n\n"
            "Нормализация данных *.csv с разных рабочих мест «Бастион»")

    def set_status(self, text, color=None):
        def _set():
            self.status_label.config(text=text, fg=color or self.COLORS['text_secondary'])
            self.root.update_idletasks()
        self._ui(_set)

    def start_progress(self):
        def _start():
            self.progress.pack(pady=(0, 10))
            self.progress.start(10)
            self.btn_process.config(state=DISABLED)
            self.btn_check.config(state=DISABLED)
            self.root.update()
        self._ui(_start)

    def stop_progress(self):
        def _stop():
            self.progress.stop()
            self.progress.pack_forget()
            self.btn_process.config(state=NORMAL)
            self.btn_check.config(state=NORMAL)
            self.root.update()
        self._ui(_stop)

    def log(self, msg, tag=None):
        def _log():
            self.log_text.config(state=NORMAL)

            # Определяем цвет по префиксу сообщения
            if tag:
                color = tag
            elif msg.startswith('✅'):
                color = 'success'
            elif msg.startswith('⚠') or msg.startswith('❗'):
                color = 'warning'
            elif msg.startswith('❌'):
                color = 'error'
            elif msg.startswith('📁') or msg.startswith('💾'):
                color = 'info'
            elif msg.startswith('📊') or msg.startswith('🏢') or msg.startswith('🔒'):
                color = 'stat'
            else:
                color = 'default'

            colors = {
                'success': '#52c41a',
                'warning': '#faad14',
                'error': '#ff4d4f',
                'info': '#1890ff',
                'stat': '#722ed1',
                'default': '#d4d4d4'
            }

            self.log_text.tag_config(color, foreground=colors.get(color, '#d4d4d4'))
            self.log_text.insert(END, msg + "\n", color)
            self.log_text.see(END)
            self.log_text.config(state=DISABLED)
            self.log_text.update_idletasks()

            if hasattr(self, 'log_file'):
                with open(self.log_file, "a", encoding="utf-8") as f:
                    f.write(msg + "\n")
        self._ui(_log)

    def detect_encoding(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                f.read(1024)
            return 'utf-8'
        except UnicodeDecodeError:
            return 'cp1251'

    def check_export_file(self):
        """Проверка готового xlsx файла на соответствие структуры"""
        file_path = filedialog.askopenfilename(
            title="Выберите файл для проверки",
            filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls")]
        )
        if not file_path:
            return

        # Создаем лог для проверки
        folder = os.path.dirname(file_path)
        timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        self.log_file = os.path.join(folder, f"Бастион_Экспорт_Проверка_{timestamp}.txt")
        open(self.log_file, "w", encoding="utf-8").close()

        self.log(f"Проверка файла: {file_path}")
        self.log("="*50)

        try:
            # Загружаем файл
            df = pd.read_excel(file_path, dtype=str, keep_default_na=False, na_filter=False)
            self.log(f"Файл успешно загружен. Строк: {len(df)}, Столбцов: {len(df.columns)}")

            issues_found = False

            # 1. Проверка на наличие всех необходимых столбцов
            missing_columns = []
            extra_columns = []
            
            for field in TARGET_FIELDS:
                if field not in df.columns:
                    missing_columns.append(field)
            
            for col in df.columns:
                if col not in TARGET_FIELDS:
                    extra_columns.append(col)

            if missing_columns:
                issues_found = True
                self.log(f"❌ ОТСУТСТВУЮЩИЕ СТОЛБЦЫ ({len(missing_columns)}):")
                for col in missing_columns:
                    self.log(f"  - {col}")

            if extra_columns:
                issues_found = True
                self.log(f"❌ ЛИШНИЕ СТОЛБЦЫ ({len(extra_columns)}):")
                for col in extra_columns:
                    self.log(f"  - {col}")

            # 2. Проверка порядка столбцов
            actual_order = list(df.columns)
            expected_order = TARGET_FIELDS.copy()
            
            if actual_order != expected_order:
                issues_found = True
                self.log("❌ ПОРЯДОК СТОЛБЦОВ НЕ СООТВЕТСТВУЕТ ТРЕБУЕМОМУ:")
                self.log("  Ожидаемый порядок:")
                for i, col in enumerate(expected_order):
                    self.log(f"    {i+1}. {col}")
                self.log("  Фактический порядок:")
                for i, col in enumerate(actual_order):
                    self.log(f"    {i+1}. {col}")

            # 3. Проверка FULLCARDCODE
            if 'FULLCARDCODE' in df.columns:
                df['FULLCARDCODE'] = df['FULLCARDCODE'].astype(str).str.strip()
                hex_pattern = re.compile(r'^[0-9A-Fa-f]{12}$')
                
                # Проверка формата
                invalid_codes = df[~df['FULLCARDCODE'].apply(lambda x: bool(hex_pattern.fullmatch(x))) & (df['FULLCARDCODE'] != '')]
                if len(invalid_codes) > 0:
                    issues_found = True
                    self.log(f"❌ НЕКОРРЕКТНЫЕ FULLCARDCODE ({len(invalid_codes)}):")
                    for idx, row in invalid_codes.head(10).iterrows():
                        self.log(f"  Строка {idx+2}: {row['FULLCARDCODE']}")

                # Проверка дубликатов FULLCARDCODE
                valid_codes_df = df[df['FULLCARDCODE'].apply(lambda x: bool(hex_pattern.fullmatch(x)))]
                duplicated_codes = valid_codes_df[valid_codes_df.duplicated(subset=['FULLCARDCODE'], keep=False)]
                
                if len(duplicated_codes) > 0:
                    issues_found = True
                    unique_duplicated = duplicated_codes['FULLCARDCODE'].nunique()
                    self.log(f"❌ ДУБЛИКАТЫ FULLCARDCODE ({len(duplicated_codes)} строк, {unique_duplicated} уникальных):")
                    for code in duplicated_codes['FULLCARDCODE'].unique()[:10]:
                        count = len(duplicated_codes[duplicated_codes['FULLCARDCODE'] == code])
                        self.log(f"  {code}: {count} раз")
                    if len(duplicated_codes['FULLCARDCODE'].unique()) > 10:
                        self.log(f"  ... и ещё {len(duplicated_codes['FULLCARDCODE'].unique()) - 10} дубликатов")

            else:
                issues_found = True
                self.log("❌ ОТСУТСТВУЕТ СТОЛБЕЦ FULLCARDCODE")

            # 4. Проверка дубликатов строк
            original_len = len(df)
            unique_df = df.drop_duplicates()
            if len(unique_df) != original_len:
                issues_found = True
                duplicate_count = original_len - len(unique_df)
                self.log(f"❌ ДУБЛИКАТЫ СТРОК ({duplicate_count})")

            # 5. Проверка обязательных полей NAME и TABLENO
            if 'NAME' in df.columns and 'TABLENO' in df.columns:
                empty_name_table = df[
                    ((df['NAME'].isna()) | (df['NAME'] == '') | (df['NAME'].str.strip() == '')) |
                    ((df['TABLENO'].isna()) | (df['TABLENO'] == '') | (df['TABLENO'].str.strip() == ''))
                ]
                
                if len(empty_name_table) > 0:
                    issues_found = True
                    self.log(f"❌ СТРОКИ БЕЗ NAME ИЛИ TABLENO ({len(empty_name_table)}):")
            else:
                issues_found = True
                self.log("❌ ОТСУТСТВУЮТ СТОЛБЦЫ NAME ИЛИ TABLENO")

            # 6. Проверка пустых строк
            all_empty_rows = df[df.astype(str).apply(lambda col: col.str.strip()).eq('').all(axis=1)]
            if len(all_empty_rows) > 0:
                issues_found = True
                self.log(f"❌ ПУСТЫЕ СТРОКИ ({len(all_empty_rows)}):")

            # 7. Проверка на пробелы в строковых данных
            string_columns = df.select_dtypes(include=['object']).columns
            rows_with_leading_trailing_spaces = 0
            
            for col in string_columns:
                if col in df.columns:
                    mask = df[col].astype(str).str.contains(r'^\s|\s$', regex=True, na=False)
                    rows_with_leading_trailing_spaces += mask.sum()
            
            if rows_with_leading_trailing_spaces > 0:
                issues_found = True
                self.log(f"❌ ДАННЫЕ С НАЧАЛЬНЫМИ/КОНЕЧНЫМИ ПРОБЕЛАМИ ({rows_with_leading_trailing_spaces})")

            if not issues_found:
                self.log("✅ Файл соответствует всем требованиям!")
                messagebox.showinfo("Проверка завершена", f"Файл {os.path.basename(file_path)} соответствует всем требованиям!")
            else:
                self.log("❌ Обнаружены проблемы в файле!")
                messagebox.showwarning("Проверка завершена", f"Файл {os.path.basename(file_path)} содержит ошибки! Подробности в логе.")

        except Exception as e:
            self.log(f"❌ ОШИБКА при проверке файла: {str(e)}")
            messagebox.showerror("Ошибка", f"Не удалось проверить файл: {str(e)}")

    def run_process(self):
        from tkinter import filedialog
        folder = filedialog.askdirectory(title="Выберите папку с CSV-файлами")
        if not folder:
            return

        self.log_file = os.path.join(folder, "export_log.txt")
        open(self.log_file, "w", encoding="utf-8").close()

        self.log("═" * 50, 'info')
        self.log(f"🚀 Начат экспорт из папки: {folder}", 'info')
        self.log("═" * 50, 'info')
        
        self.start_progress()
        self.set_status("Обработка файлов...", self.COLORS['primary'])
        
        # Запускаем обработку в отдельном потоке
        thread = threading.Thread(target=self._run_process_thread, args=(folder,))
        thread.daemon = True
        thread.start()

    def _run_process_thread(self, folder):
        csv_files = glob.glob(os.path.join(folder, "*.csv"))
        if not csv_files:
            self.stop_progress()
            self.set_status("Ошибка: файлы не найдены", self.COLORS['error'])
            self._ui(messagebox.showerror, "❌ Ошибка", "В папке нет CSV-файлов!")
            self.log("ОШИБКА: CSV-файлы не найдены.", 'error')
            return

        self.log(f"Найдено {len(csv_files)} CSV-файлов. Загрузка...")

        all_dfs = []
        for f in csv_files:
            try:
                enc = self.detect_encoding(f)
                df = pd.read_csv(f, sep=';', quotechar='"', encoding=enc,
                                 dtype=str, keep_default_na=False, na_filter=False)
                all_dfs.append(df)
                self.log(f" + {os.path.basename(f)} — {len(df)} строк (кодировка: {enc})")
            except Exception as e:
                self.log(f" ОШИБКА при чтении {f}: {str(e)}")

        if not all_dfs:
            self.stop_progress()
            self.set_status("Ошибка: файлы не загружены", self.COLORS['error'])
            self.log("❌ ОШИБКА: ни один файл не загружен.", 'error')
            return

        combined = pd.concat(all_dfs, ignore_index=True)
        initial_count = len(combined)
        self.log(f"\nВсего строк после объединения: {initial_count}")

        # Удаление заглушек
        placeholder_cols = ['NAME', 'FIRSTNAME', 'SECONDNAME']
        if all(col in combined.columns for col in placeholder_cols):
            mask_bad = (combined['NAME'] == 'Фамилия') & (combined['FIRSTNAME'] == 'Имя') & (combined['SECONDNAME'] == 'Отчество')
            bad_rows = mask_bad.sum()
            combined = combined[~mask_bad].copy()
            self.log(f"\nУдалено полей с русскими названиями: {bad_rows}")
        else:
            self.log("\n⚠️ Пропущено удаление заглушек: отсутствуют столбцы NAME/FIRSTNAME/SECONDNAME")

        # Применяем strip ко всем строковым значениям
        for col in combined.columns:
            if combined[col].dtype == 'object':
                combined[col] = combined[col].astype(str).str.strip()
        
        self.log("✅ Удалены начальные и конечные пробелы из всех строковых полей")

        # Валидация FULLCARDCODE — сохраняем ВСЕХ удалённых
        if 'FULLCARDCODE' in combined.columns:
            combined['FULLCARDCODE'] = combined['FULLCARDCODE'].astype(str).str.strip()
            hex_pattern = re.compile(r'^[0-9A-Fa-f]{12}$')
            valid_mask = combined['FULLCARDCODE'].apply(lambda x: bool(hex_pattern.fullmatch(x)))
            invalid_count = (~valid_mask).sum()

            if invalid_count > 0:
                # Выделяем ВСЕХ, кого удалим
                rejected_df = combined[~valid_mask].copy()
                rejected_file = os.path.join(folder, "rejected_FULLCARDCODE.xlsx")
                rejected_df.to_excel(rejected_file, sheet_name='Отклонённые', index=False)
                self.log(f"⚠️ УДАЛЕНО строк с битым FULLCARDCODE: {invalid_count}")
                self.log(f"📁 Полный список сохранён в: {rejected_file}")
                # Оставляем только валидные
                combined = combined[valid_mask].copy()
            else:
                self.log("✅ FULLCARDCODE: все значения корректны")
        else:
            self.log("❌ ОШИБКА: отсутствует поле FULLCARDCODE — все строки отклонены")
            combined = combined.iloc[0:0]

        # Пустые строки
        before_empty = len(combined)
        empty_mask = combined.astype(str).apply(lambda col: col.str.strip()).eq('').all(axis=1)
        combined = combined[~empty_mask].copy()
        self.log(f"Удалено пустых строк: {before_empty - len(combined)}")

        # NAME / TABLENO — сохраняем отклонённых
        required_cols = ['NAME', 'TABLENO']
        missing_required_cols = [col for col in required_cols if col not in combined.columns]
        if missing_required_cols:
            rejected_count = len(combined)
            if rejected_count > 0:
                rejected_file = os.path.join(folder, "rejected_NAME_TABLENO.xlsx")
                combined.to_excel(rejected_file, sheet_name='Отклонённые', index=False)
                self.log(f"⚠️ ОТСУТСТВУЮТ обязательные столбцы: {', '.join(missing_required_cols)}")
                self.log(f"⚠️ УДАЛЕНО строк без возможности проверки NAME/TABLENO: {rejected_count}")
                self.log(f"📁 Полный список сохранён в: {rejected_file}")
            combined = combined.iloc[0:0]
        else:
            req_values = combined[required_cols].astype(str).apply(lambda col: col.str.strip())
            required_mask = (req_values != '').all(axis=1)
            rejected_req = combined[~required_mask].copy()
            rejected_count = len(rejected_req)

            if rejected_count > 0:
                rejected_file = os.path.join(folder, "rejected_NAME_TABLENO.xlsx")
                rejected_req.to_excel(rejected_file, sheet_name='Отклонённые', index=False)
                self.log(f"⚠️ УДАЛЕНО строк без NAME/TABLENO: {rejected_count}")
                self.log(f"📁 Полный список сохранён в: {rejected_file}")
            else:
                self.log("✅ Все строки содержат NAME и TABLENO")

            combined = combined[required_mask].copy()

        # Проверка наличия должности (POST)
        if 'POST' in combined.columns:
            # Удаляем строки, где POST пустой (после strip)
            post_mask = combined['POST'].astype(str).str.strip() != ''
            rejected_no_post = combined[~post_mask].copy()
            rejected_no_post_count = len(rejected_no_post)

            if rejected_no_post_count > 0:
                rejected_post_file = os.path.join(folder, "rejected_no_POST.xlsx")
                rejected_no_post.to_excel(rejected_post_file, sheet_name='Без должности', index=False)
                self.log(f"⚠️ УДАЛЕНО строк без должности (POST): {rejected_no_post_count}")
                self.log(f"📁 Список сохранён в: {rejected_post_file}")
            else:
                self.log("✅ Все строки содержат должность (POST)")

            # Оставляем только строки с непустым POST
            combined = combined[post_mask].copy()
        else:
            # Если столбца POST вообще нет — считаем, что все строки без должности
            rejected_no_post_count = len(combined)
            if rejected_no_post_count > 0:
                rejected_post_file = os.path.join(folder, "rejected_no_POST.xlsx")
                combined.to_excel(rejected_post_file, sheet_name='Без должности', index=False)
                self.log(f"⚠️ СТОЛБЕЦ POST ОТСУТСТВУЕТ — все {rejected_no_post_count} строк отклонены")
                self.log(f"📁 Список сохранён в: {rejected_post_file}")
                combined = combined.iloc[0:0]  # Очищаем DataFrame
            else:
                self.log("✅ Нет данных для обработки (POST отсутствует, но и строк нет)")        

        # Проверка дубликатов по FULLCARDCODE
        if 'FULLCARDCODE' in combined.columns:
            # Находим дубликаты по FULLCARDCODE
            duplicated_mask = combined.duplicated(subset=['FULLCARDCODE'], keep=False)
            duplicated_count = duplicated_mask.sum()
            
            if duplicated_count > 0:
                duplicated_df = combined[duplicated_mask].copy()
                duplicated_file = os.path.join(folder, "duplicated_FULLCARDCODE.xlsx")
                duplicated_df.to_excel(duplicated_file, sheet_name='Дубликаты', index=False)
                
                # Получаем уникальные дублирующиеся коды
                unique_duplicated_codes = duplicated_df['FULLCARDCODE'].unique()
                self.log(f"⚠️ НАЙДЕНО дубликатов по FULLCARDCODE: {duplicated_count} строк")
                self.log(f"⚠️ Уникальных дублирующихся кодов: {len(unique_duplicated_codes)}")
                self.log(f"📁 Дубликаты сохранены в: {duplicated_file}")
                
                # Удаляем дубликаты, оставляя первый экземпляр
                combined = combined.drop_duplicates(subset=['FULLCARDCODE'], keep='first')
                self.log(f"✅ После удаления дубликатов: {len(combined)} строк")
            else:
                self.log("✅ Нет дубликатов по FULLCARDCODE")

        # Дубликаты по всем полям (после удаления дубликатов по FULLCARDCODE)
        before_dupes = len(combined)
        combined.drop_duplicates(inplace=True)
        self.log(f"Удалено дубликатов по всем полям: {before_dupes - len(combined)}")

        # WORG6 → WORG7
        if all(col in combined.columns for col in ['WORG6','WORG7','WORG8']):
            mask_fix = (combined['WORG7'].str.strip() == '') & (combined['WORG8'].str.strip() == '') & (combined['WORG6'].str.strip() != '')
            fixed = mask_fix.sum()
            if fixed:
                combined.loc[mask_fix, 'WORG7'] = combined.loc[mask_fix, 'WORG6']
                self.log(f"Перенос названия организации из WORG6 → WORG7: {fixed}")

        # WDEP8
        if 'WDEP8' not in combined.columns:
            combined['WDEP8'] = 'Нет данных'
        else:
            mask_empty = combined['WDEP8'].str.strip() == ''
            combined.loc[mask_empty, 'WDEP8'] = 'Нет данных'
            self.log(f"Заполнено пустых *Подразделений*: {mask_empty.sum()}")

        # Статистика по отделам
        if 'WDEP8' in combined.columns:
            dep_stats = combined['WDEP8'].value_counts()
            self.log("\n📊 Статистика по отделам (топ-10):")
            for i, (dep, count) in enumerate(dep_stats.head(10).items()):
                self.log(f"   {i+1}. {dep}: {count} человек")
            
            if len(dep_stats) > 10:
                self.log(f"   ... и ещё {len(dep_stats) - 10} отделов")

        # Статистика по организациям
        org_columns = [col for col in ['WORG1', 'WORG2', 'WORG3', 'WORG4', 'WORG5', 'WORG6', 'WORG7', 'WORG8'] if col in combined.columns]
        if org_columns:
            # Используем WORG7 как основной источник информации об организации
            if 'WORG7' in combined.columns and combined['WORG7'].notna().any():
                org_stats = combined['WORG7'].value_counts()
                self.log("\n🏢 Статистика по организациям (топ-10):")
                for i, (org, count) in enumerate(org_stats.head(10).items()):
                    if org and org.strip() != '':
                        self.log(f"   {i+1}. {org}: {count} человек")
                
                if len(org_stats) > 10:
                    self.log(f"   ... и ещё {len(org_stats) - 10} организаций")
            elif org_columns:
                # Если WORG7 пустой, используем любое из WORG полей
                org_data = pd.Series(dtype=str)
                for col in org_columns:
                    org_data = pd.concat([org_data, combined[col]])
                org_stats = org_data.value_counts()
                
                self.log("\n🏢 Статистика по организациям (топ-10):")
                count = 0
                for org, org_count in org_stats.head(10).items():
                    if org and org.strip() != '':
                        self.log(f"   {count+1}. {org}: {org_count} человек")
                        count += 1
                
                if len([x for x in org_stats.head(10).items() if x[0] and x[0].strip() != '']) < 10:
                    remaining_orgs = len([x for x in org_stats.items() if x[0] and x[0].strip() != '']) - 10
                    if remaining_orgs > 0:
                        self.log(f"   ... и ещё {remaining_orgs} организаций")

        # Статистика по заблокированным пропускам
        if 'IS_BLOCKED' in combined.columns:
            blocked_count = (combined['IS_BLOCKED'] == '1').sum()
            total_count = len(combined)
            if total_count > 0:
                blocked_percent = blocked_count / total_count * 100
                self.log(f"\n🔒 Статистика по заблокированным пропускам: {blocked_count} из {total_count} ({blocked_percent:.2f}%)")
            else:
                self.log("\n🔒 Статистика по заблокированным пропускам: 0 из 0 (0.00%)")

        self.log("\n💾 Ждем сохранение файла")

        for col in TARGET_FIELDS:
            if col not in combined.columns:
                combined[col] = ''

        combined = combined[TARGET_FIELDS]

        timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        output_file = os.path.join(folder, f"Бастион_Экспорт_{timestamp}.xlsx")
        combined.to_excel(output_file, sheet_name='Лист1', index=False)
        self.log(f"\nФайл сохранён: {output_file}")

        # Пересохраняем через Excel COM
        if HAS_WIN32:
            excel = None
            wb = None
            try:
                excel = win32.Dispatch("Excel.Application")
                excel.Visible = False
                wb = excel.Workbooks.Open(output_file)
                wb.Save()
                self.log("✅ Файл пересохранён через Excel (структура выровнена)")
            except Exception as e:
                self.log(f"⚠ Не удалось пересохранить через Excel: {str(e)}")
            finally:
                try:
                    if wb is not None:
                        wb.Close(SaveChanges=True)
                finally:
                    if excel is not None:
                        excel.Quit()
        else:
            self.log("⚠ Модуль win32com не установлен — пересохранение пропущено")

        self.stop_progress()
        self.set_status("Готово! Обработано записей: " + str(len(combined)), self.COLORS['success'])
        self._ui(messagebox.showinfo, "✅ Готово!", f"Экспорт завершён!\n\n📁 Файл: {output_file}\n📝 Лог: export_log.txt\n📊 Обработано: {len(combined)} записей")

def main():
    root = Tk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
