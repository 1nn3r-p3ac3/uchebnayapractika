# salary_system_desktop_with_edit.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import sqlite3
import csv
import json
from datetime import datetime
import os

# ============================================================================
# КЛАССЫ ДЛЯ РАСЧЕТОВ
# ============================================================================

class Employee:
    def __init__(self, **kwargs):
        self.id = kwargs.get('id')
        self.full_name = kwargs.get('full_name', '')
        self.position = kwargs.get('position', '')
        self.base_salary = float(kwargs.get('base_salary', 0))
        self.bank_account = kwargs.get('bank_account', '')
        self.tax_id = kwargs.get('tax_id', '')
        self.hire_date = kwargs.get('hire_date', '')
        self.department = kwargs.get('department', '')
        self.email = kwargs.get('email', '')
        self.phone = kwargs.get('phone', '')
        self.is_active = kwargs.get('is_active', True)
    
    def to_dict(self):
        return self.__dict__.copy()

class SalaryCalculator:
    def __init__(self, tax_rate=0.13, overtime_rate=1.5):
        self.tax_rate = tax_rate
        self.overtime_rate = overtime_rate
    
    def calculate_base_salary(self, employee, worked_days, total_days):
        daily_salary = employee.base_salary / total_days
        return daily_salary * worked_days
    
    def calculate_overtime(self, overtime_hours, hourly_rate):
        return overtime_hours * hourly_rate * self.overtime_rate
    
    def calculate_bonus(self, employee, kpi_score):
        bonus_rates = {'A': 0.3, 'B': 0.15, 'C': 0.05, 'D': 0.0}
        return employee.base_salary * bonus_rates.get(kpi_score, 0)
    
    def calculate_total_income(self, base_salary, bonus, overtime_pay, sick_pay=0, vacation_pay=0):
        return base_salary + bonus + overtime_pay + sick_pay + vacation_pay

class TaxService:
    def __init__(self, ndfl_rate=0.13):
        self.ndfl_rate = ndfl_rate
    
    def calculate_ndfl(self, total_income):
        return total_income * self.ndfl_rate
    
    def calculate_net_salary(self, total_income, tax_amount):
        return total_income - tax_amount

# ============================================================================
# ГЛАВНОЕ ОКНО ПРИЛОЖЕНИЯ
# ============================================================================

class SalarySystemApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Система расчета заработной платы")
        self.root.geometry("1200x700")
        
        # Инициализация компонентов
        self.db_path = "salary_system.db"
        self.calculator = SalaryCalculator()
        self.tax_service = TaxService()
        
        # Создание вкладок
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Создание вкладок
        self.create_dashboard_tab()
        self.create_employees_tab()
        self.create_payroll_tab()
        self.create_reports_tab()
        self.create_import_tab()
        
        # Загрузка данных при запуске
        self.load_employees()
        
    def create_dashboard_tab(self):
        """Создание вкладки Дашборд"""
        dashboard_frame = ttk.Frame(self.notebook)
        self.notebook.add(dashboard_frame, text='Дашборд')
        
        # Заголовок
        title_label = ttk.Label(dashboard_frame, 
                               text="Система расчета заработной платы",
                               font=('Arial', 16, 'bold'))
        title_label.pack(pady=20)
        
        # Статистика
        stats_frame = ttk.LabelFrame(dashboard_frame, text="Статистика", padding=20)
        stats_frame.pack(fill='x', padx=20, pady=10)
        
        self.total_employees_label = ttk.Label(stats_frame, text="Всего сотрудников: 0")
        self.total_employees_label.grid(row=0, column=0, padx=20, pady=5, sticky='w')
        
        self.active_employees_label = ttk.Label(stats_frame, text="Активных: 0")
        self.active_employees_label.grid(row=1, column=0, padx=20, pady=5, sticky='w')
        
        self.total_payroll_label = ttk.Label(stats_frame, text="Общий фонд оплаты: 0 ₽")
        self.total_payroll_label.grid(row=0, column=1, padx=20, pady=5, sticky='w')
        
        self.avg_salary_label = ttk.Label(stats_frame, text="Средняя зарплата: 0 ₽")
        self.avg_salary_label.grid(row=1, column=1, padx=20, pady=5, sticky='w')
        
        # Быстрые действия
        actions_frame = ttk.LabelFrame(dashboard_frame, text="Быстрые действия", padding=20)
        actions_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Button(actions_frame, text="Добавить сотрудника", 
                  command=self.show_add_employee_dialog).grid(row=0, column=0, padx=10, pady=5)
        ttk.Button(actions_frame, text="Рассчитать зарплату", 
                  command=lambda: self.notebook.select(2)).grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(actions_frame, text="Импорт из 1С", 
                  command=lambda: self.notebook.select(4)).grid(row=0, column=2, padx=10, pady=5)
        
        # Информация о системе
        info_frame = ttk.LabelFrame(dashboard_frame, text="Информация о системе", padding=20)
        info_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Label(info_frame, text="Версия: 1.0.0").pack(anchor='w')
        ttk.Label(info_frame, text="База данных: salary_system.db").pack(anchor='w')
        ttk.Label(info_frame, text=f"Дата: {datetime.now().strftime('%d.%m.%Y')}").pack(anchor='w')
        
    def create_employees_tab(self):
        """Создание вкладки Сотрудники"""
        employees_frame = ttk.Frame(self.notebook)
        self.notebook.add(employees_frame, text='Сотрудники')
        
        # Панель управления
        control_frame = ttk.Frame(employees_frame)
        control_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(control_frame, text="Добавить", 
                  command=self.show_add_employee_dialog).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Редактировать", 
                  command=self.show_edit_employee_dialog).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Удалить", 
                  command=self.delete_employee).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Обновить", 
                  command=self.load_employees).pack(side='right', padx=5)
        
        # Поиск
        search_frame = ttk.Frame(employees_frame)
        search_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(search_frame, text="Поиск:").pack(side='left', padx=5)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side='left', padx=5)
        search_entry.bind('<KeyRelease>', self.search_employees)
        
        # Таблица сотрудников
        columns = ('ID', 'ФИО', 'Должность', 'Отдел', 'Оклад', 'Статус')
        self.employees_tree = ttk.Treeview(employees_frame, columns=columns, show='headings', height=20)
        
        # Настройка колонок
        for col in columns:
            self.employees_tree.heading(col, text=col)
            self.employees_tree.column(col, width=100)
        
        self.employees_tree.column('ФИО', width=200)
        self.employees_tree.column('Должность', width=150)
        
        # Скроллбар
        scrollbar = ttk.Scrollbar(employees_frame, orient='vertical', command=self.employees_tree.yview)
        self.employees_tree.configure(yscrollcommand=scrollbar.set)
        
        self.employees_tree.pack(side='left', fill='both', expand=True, padx=(10, 0), pady=10)
        scrollbar.pack(side='right', fill='y', padx=(0, 10), pady=10)
        
        # Привязка двойного клика
        self.employees_tree.bind('<Double-Button-1>', self.view_employee_details)
        
    def create_payroll_tab(self):
        """Создание вкладки Расчет ЗП"""
        payroll_frame = ttk.Frame(self.notebook)
        self.notebook.add(payroll_frame, text='Расчет ЗП')
        
        # Левая панель - параметры расчета
        left_frame = ttk.Frame(payroll_frame)
        left_frame.pack(side='left', fill='y', padx=10, pady=10)
        
        # Период расчета
        period_frame = ttk.LabelFrame(left_frame, text="Период расчета", padding=10)
        period_frame.pack(fill='x', pady=5)
        
        ttk.Label(period_frame, text="Месяц:").grid(row=0, column=0, sticky='w', pady=5)
        self.month_var = tk.StringVar(value=datetime.now().strftime('%Y-%m'))
        ttk.Entry(period_frame, textvariable=self.month_var, width=15).grid(row=0, column=1, pady=5)
        
        # Параметры расчета
        params_frame = ttk.LabelFrame(left_frame, text="Параметры расчета", padding=10)
        params_frame.pack(fill='x', pady=5)
        
        ttk.Label(params_frame, text="Рабочих дней:").grid(row=0, column=0, sticky='w', pady=2)
        self.working_days_var = tk.StringVar(value="22")
        ttk.Entry(params_frame, textvariable=self.working_days_var, width=10).grid(row=0, column=1, pady=2)
        
        ttk.Label(params_frame, text="Отработано:").grid(row=1, column=0, sticky='w', pady=2)
        self.worked_days_var = tk.StringVar(value="20")
        ttk.Entry(params_frame, textvariable=self.worked_days_var, width=10).grid(row=1, column=1, pady=2)
        
        ttk.Label(params_frame, text="Коэф. премии:").grid(row=2, column=0, sticky='w', pady=2)
        self.bonus_kpi_var = tk.StringVar(value="B")
        bonus_combo = ttk.Combobox(params_frame, textvariable=self.bonus_kpi_var, 
                                  values=['A', 'B', 'C', 'D'], width=8)
        bonus_combo.grid(row=2, column=1, pady=2)
        
        ttk.Label(params_frame, text="Сверхурочные часы:").grid(row=3, column=0, sticky='w', pady=2)
        self.overtime_var = tk.StringVar(value="5")
        ttk.Entry(params_frame, textvariable=self.overtime_var, width=10).grid(row=3, column=1, pady=2)
        
        # Кнопки расчета
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(fill='x', pady=20)
        
        ttk.Button(button_frame, text="Рассчитать для всех", 
                  command=self.calculate_all_payroll).pack(pady=5)
        ttk.Button(button_frame, text="Рассчитать выбранных", 
                  command=self.calculate_selected_payroll).pack(pady=5)
        ttk.Button(button_frame, text="Экспорт в CSV", 
                  command=self.export_payroll_csv).pack(pady=5)
        
        # Правая панель - результаты расчета
        right_frame = ttk.Frame(payroll_frame)
        right_frame.pack(side='right', fill='both', expand=True, padx=10, pady=10)
        
        # Таблица результатов
        columns = ('ФИО', 'Оклад', 'Премия', 'Сверхурочные', 'Начислено', 'НДФЛ', 'К выплате')
        self.payroll_tree = ttk.Treeview(right_frame, columns=columns, show='headings', height=20)
        
        for col in columns:
            self.payroll_tree.heading(col, text=col)
            self.payroll_tree.column(col, width=100)
        
        self.payroll_tree.column('ФИО', width=150)
        
        # Скроллбар
        scrollbar = ttk.Scrollbar(right_frame, orient='vertical', command=self.payroll_tree.yview)
        self.payroll_tree.configure(yscrollcommand=scrollbar.set)
        
        self.payroll_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Итоги
        summary_frame = ttk.LabelFrame(right_frame, text="Итоги", padding=10)
        summary_frame.pack(fill='x', pady=10)
        
        self.total_income_label = ttk.Label(summary_frame, text="Общая сумма: 0 ₽")
        self.total_income_label.pack(side='left', padx=20)
        
        self.total_tax_label = ttk.Label(summary_frame, text="Общий НДФЛ: 0 ₽")
        self.total_tax_label.pack(side='left', padx=20)
        
        self.total_net_label = ttk.Label(summary_frame, text="К выплате: 0 ₽")
        self.total_net_label.pack(side='left', padx=20)
        
    def create_reports_tab(self):
        """Создание вкладки Отчеты"""
        reports_frame = ttk.Frame(self.notebook)
        self.notebook.add(reports_frame, text='Отчеты')
        
        # Виджеты для отчетов
        ttk.Label(reports_frame, text="Генерация отчетов", 
                 font=('Arial', 14, 'bold')).pack(pady=20)
        
        reports_list_frame = ttk.Frame(reports_frame)
        reports_list_frame.pack(pady=10)
        
        reports = [
            ("Расчетная ведомость", "payroll_report"),
            ("Список сотрудников", "employee_list"),
            ("Налоговый отчет", "tax_report"),
            ("Отчет по отделам", "department_report")
        ]
        
        for report_name, report_code in reports:
            frame = ttk.Frame(reports_list_frame)
            frame.pack(fill='x', pady=5, padx=50)
            
            ttk.Label(frame, text=report_name, width=20).pack(side='left')
            ttk.Button(frame, text="Сгенерировать", 
                      command=lambda rc=report_code: self.generate_report(rc)).pack(side='right')
        
        # Текстовое поле для предпросмотра
        ttk.Label(reports_frame, text="Предпросмотр отчета:").pack(pady=10)
        
        self.report_text = scrolledtext.ScrolledText(reports_frame, height=15, width=100)
        self.report_text.pack(padx=20, pady=10)
        
        # Кнопки управления отчетом
        button_frame = ttk.Frame(reports_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Сохранить как CSV", 
                  command=self.save_report_csv).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Сохранить как TXT", 
                  command=self.save_report_txt).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Очистить", 
                  command=lambda: self.report_text.delete(1.0, tk.END)).pack(side='left', padx=5)
        
    def create_import_tab(self):
        """Создание вкладки Импорт"""
        import_frame = ttk.Frame(self.notebook)
        self.notebook.add(import_frame, text='Импорт')
        
        ttk.Label(import_frame, text="Импорт данных из 1С", 
                 font=('Arial', 14, 'bold')).pack(pady=20)
        
        # Формат импорта
        format_frame = ttk.LabelFrame(import_frame, text="Формат файла", padding=10)
        format_frame.pack(pady=10, padx=20, fill='x')
        
        self.import_format = tk.StringVar(value="csv")
        ttk.Radiobutton(format_frame, text="CSV (разделитель ;)", 
                       variable=self.import_format, value="csv").pack(anchor='w')
        ttk.Radiobutton(format_frame, text="JSON", 
                       variable=self.import_format, value="json").pack(anchor='w')
        
        # Выбор файла
        file_frame = ttk.Frame(import_frame)
        file_frame.pack(pady=10, padx=20, fill='x')
        
        ttk.Label(file_frame, text="Файл:").pack(side='left')
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=50).pack(side='left', padx=10)
        ttk.Button(file_frame, text="Обзор...", 
                  command=self.browse_import_file).pack(side='left')
        
        # Параметры импорта
        params_frame = ttk.LabelFrame(import_frame, text="Параметры импорта", padding=10)
        params_frame.pack(pady=10, padx=20, fill='x')
        
        self.update_existing_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(params_frame, text="Обновлять существующих сотрудников", 
                       variable=self.update_existing_var).pack(anchor='w')
        
        self.create_missing_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(params_frame, text="Создавать отсутствующих сотрудников", 
                       variable=self.create_missing_var).pack(anchor='w')
        
        # Кнопка импорта
        ttk.Button(import_frame, text="Импортировать данные", 
                  command=self.import_from_1c).pack(pady=20)
        
        # Область предпросмотра
        preview_frame = ttk.LabelFrame(import_frame, text="Предпросмотр данных", padding=10)
        preview_frame.pack(pady=10, padx=20, fill='both', expand=True)
        
        self.import_text = scrolledtext.ScrolledText(preview_frame, height=15)
        self.import_text.pack(fill='both', expand=True)
        
    # ============================================================================
    # МЕТОДЫ РАБОТЫ С БАЗОЙ ДАННЫХ
    # ============================================================================
    
    def get_connection(self):
        """Получение соединения с базой данных"""
        try:
            return sqlite3.connect(self.db_path)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось подключиться к базе данных:\n{e}")
            return None
    
    def load_employees(self):
        """Загрузка списка сотрудников из базы данных"""
        try:
            conn = self.get_connection()
            if not conn:
                return
            
            cursor = conn.cursor()
            
            # Проверяем существование таблицы
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='employees'")
            if not cursor.fetchone():
                messagebox.showwarning("Внимание", "Таблица employees не найдена в базе данных")
                conn.close()
                return
            
            # Получаем всех сотрудников
            cursor.execute("SELECT id, full_name, position, department, base_salary, is_active FROM employees")
            employees = cursor.fetchall()
            conn.close()
            
            # Очищаем таблицу
            for item in self.employees_tree.get_children():
                self.employees_tree.delete(item)
            
            # Заполняем таблицу
            for emp in employees:
                status = "Активен" if emp[5] else "Неактивен"
                self.employees_tree.insert('', 'end', values=(
                    emp[0], emp[1], emp[2], emp[3], f"{emp[4]:,.2f} ₽", status
                ))
            
            # Обновляем статистику на дашборде
            self.update_dashboard_stats(employees)
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить сотрудников:\n{e}")
    
    def update_dashboard_stats(self, employees):
        """Обновление статистики на дашборде"""
        try:
            total = len(employees)
            active = sum(1 for emp in employees if emp[5])
            total_salary = sum(emp[4] for emp in employees)
            avg_salary = total_salary / total if total > 0 else 0
            
            self.total_employees_label.config(text=f"Всего сотрудников: {total}")
            self.active_employees_label.config(text=f"Активных: {active}")
            self.total_payroll_label.config(text=f"Общий фонд оплаты: {total_salary:,.2f} ₽")
            self.avg_salary_label.config(text=f"Средняя зарплата: {avg_salary:,.2f} ₽")
            
        except Exception as e:
            print(f"Ошибка обновления статистики: {e}")
    
    def search_employees(self, event=None):
        """Поиск сотрудников"""
        query = self.search_var.get().lower()
        
        # Показываем всех сотрудников если запрос пустой
        if not query:
            for item in self.employees_tree.get_children():
                self.employees_tree.item(item, tags=())
            return
        
        # Скрываем несовпадающие строки
        for item in self.employees_tree.get_children():
            values = self.employees_tree.item(item, 'values')
            if values:
                # Ищем по всем видимым полям
                text = ' '.join(str(v).lower() for v in values)
                if query in text:
                    self.employees_tree.item(item, tags=())
                else:
                    self.employees_tree.item(item, tags=('hidden',))
        
        self.employees_tree.tag_configure('hidden', foreground='gray')
    
    # ============================================================================
    # ДИАЛОГОВЫЕ ОКНА (ДОБАВЛЕНИЕ И РЕДАКТИРОВАНИЕ)
    # ============================================================================
    
    def show_add_employee_dialog(self):
        """Показать диалог добавления сотрудника"""
        self.show_employee_dialog(None, "Добавление сотрудника")
    
    def show_edit_employee_dialog(self):
        """Показать диалог редактирования сотрудника"""
        selection = self.employees_tree.selection()
        if not selection:
            messagebox.showwarning("Внимание", "Выберите сотрудника для редактирования")
            return
        
        item = self.employees_tree.item(selection[0])
        employee_id = item['values'][0]
        
        # Загружаем данные сотрудника
        employee_data = self.get_employee_by_id(employee_id)
        if not employee_data:
            messagebox.showerror("Ошибка", "Не удалось загрузить данные сотрудника")
            return
        
        self.show_employee_dialog(employee_data, "Редактирование сотрудника")
    
    def show_employee_dialog(self, employee_data, title):
        """Показать диалог добавления/редактирования сотрудника"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("500x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Сохраняем ID сотрудника для редактирования
        if employee_data:
            self.current_edit_id = employee_data['id']
        else:
            self.current_edit_id = None
        
        # Поля формы
        fields = [
            ("ФИО", "full_name"),
            ("Должность", "position"),
            ("Оклад", "base_salary"),
            ("Отдел", "department"),
            ("Банковский счет", "bank_account"),
            ("ИНН", "tax_id"),
            ("Email", "email"),
            ("Телефон", "phone"),
            ("Дата приема (ГГГГ-ММ-ДД)", "hire_date"),
            ("Статус", "is_active")
        ]
        
        entries = {}
        
        for i, (label, key) in enumerate(fields):
            ttk.Label(dialog, text=f"{label}:").grid(row=i, column=0, sticky='w', padx=10, pady=5)
            
            if key == 'base_salary':
                entry = ttk.Entry(dialog, width=30)
                if employee_data:
                    entry.insert(0, str(employee_data.get(key, 0)))
                else:
                    entry.insert(0, "0")
            
            elif key == 'hire_date':
                entry = ttk.Entry(dialog, width=30)
                if employee_data:
                    entry.insert(0, employee_data.get(key, ''))
                else:
                    entry.insert(0, datetime.now().strftime('%Y-%m-%d'))
            
            elif key == 'is_active':
                # Используем Checkbutton для статуса
                is_active_var = tk.BooleanVar(value=employee_data.get(key, True) if employee_data else True)
                entry = ttk.Checkbutton(dialog, variable=is_active_var, text="Активен")
                entries[key] = is_active_var  # Сохраняем как переменную, а не виджет
                entry.grid(row=i, column=1, padx=10, pady=5, sticky='w')
                continue  # Пропускаем добавление в entries как Entry
            
            else:
                entry = ttk.Entry(dialog, width=30)
                if employee_data:
                    entry.insert(0, employee_data.get(key, ''))
            
            entry.grid(row=i, column=1, padx=10, pady=5, sticky='ew')
            entries[key] = entry
        
        # Фрейм для кнопок
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=len(fields), column=0, columnspan=2, pady=20)
        
        if employee_data:
            ttk.Button(button_frame, text="Сохранить изменения", 
                      command=lambda: self.update_employee(entries, dialog)).pack(side='left', padx=10)
        else:
            ttk.Button(button_frame, text="Сохранить", 
                      command=lambda: self.save_employee(entries, dialog)).pack(side='left', padx=10)
        
        ttk.Button(button_frame, text="Отмена", 
                  command=dialog.destroy).pack(side='left', padx=10)
        
        dialog.columnconfigure(1, weight=1)
    
    def get_employee_by_id(self, employee_id):
        """Получить данные сотрудника по ID"""
        try:
            conn = self.get_connection()
            if not conn:
                return None
            
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM employees WHERE id = ?", (employee_id,))
            emp = cursor.fetchone()
            conn.close()
            
            if not emp:
                return None
            
            # Получаем названия колонок
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(employees)")
            columns = [col[1] for col in cursor.fetchall()]
            conn.close()
            
            # Создаем словарь с данными
            employee_data = {'id': employee_id}
            for i, col in enumerate(columns):
                if i < len(emp):
                    employee_data[col] = emp[i]
            
            return employee_data
            
        except Exception as e:
            print(f"Ошибка получения данных сотрудника: {e}")
            return None
    
    def save_employee(self, entries, dialog):
        """Сохранение нового сотрудника в базу данных"""
        try:
            conn = self.get_connection()
            if not conn:
                return
            
            cursor = conn.cursor()
            
            # Проверяем существование таблицы
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='employees'")
            if not cursor.fetchone():
                # Создаем таблицу если она не существует
                cursor.execute('''
                    CREATE TABLE employees (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        full_name TEXT NOT NULL,
                        position TEXT NOT NULL,
                        base_salary REAL NOT NULL,
                        department TEXT,
                        bank_account TEXT,
                        tax_id TEXT,
                        email TEXT,
                        phone TEXT,
                        hire_date TEXT,
                        is_active BOOLEAN DEFAULT TRUE,
                        created_at TEXT DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
            
            # Получаем значения из полей
            values = {}
            for key, entry in entries.items():
                if key == 'is_active':
                    values[key] = entry.get()  # Для BooleanVar
                else:
                    value = entry.get().strip()
                    if key == 'base_salary':
                        value = float(value) if value else 0
                    values[key] = value
            
            # Вставляем запись
            cursor.execute('''
                INSERT INTO employees (full_name, position, base_salary, department, 
                                      bank_account, tax_id, email, phone, hire_date, is_active)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                values['full_name'],
                values['position'],
                values['base_salary'],
                values['department'],
                values.get('bank_account', ''),
                values.get('tax_id', ''),
                values.get('email', ''),
                values.get('phone', ''),
                values.get('hire_date', ''),
                values.get('is_active', True)
            ))
            
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Успех", "Сотрудник успешно добавлен")
            dialog.destroy()
            self.load_employees()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить сотрудника:\n{e}")
    
    def update_employee(self, entries, dialog):
        """Обновление данных сотрудника в базе данных"""
        try:
            if not self.current_edit_id:
                messagebox.showerror("Ошибка", "Не выбран сотрудник для редактирования")
                return
            
            conn = self.get_connection()
            if not conn:
                return
            
            cursor = conn.cursor()
            
            # Получаем значения из полей
            values = {}
            for key, entry in entries.items():
                if key == 'is_active':
                    values[key] = entry.get()  # Для BooleanVar
                else:
                    value = entry.get().strip()
                    if key == 'base_salary':
                        value = float(value) if value else 0
                    values[key] = value
            
            # Обновляем запись
            cursor.execute('''
                UPDATE employees 
                SET full_name = ?, position = ?, base_salary = ?, department = ?, 
                    bank_account = ?, tax_id = ?, email = ?, phone = ?, 
                    hire_date = ?, is_active = ?
                WHERE id = ?
            ''', (
                values['full_name'],
                values['position'],
                values['base_salary'],
                values['department'],
                values.get('bank_account', ''),
                values.get('tax_id', ''),
                values.get('email', ''),
                values.get('phone', ''),
                values.get('hire_date', ''),
                values.get('is_active', True),
                self.current_edit_id
            ))
            
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Успех", "Данные сотрудника успешно обновлены")
            dialog.destroy()
            self.load_employees()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обновить данные сотрудника:\n{e}")
    
    def view_employee_details(self, event):
        """Просмотр детальной информации о сотруднике"""
        selection = self.employees_tree.selection()
        if not selection:
            return
        
        item = self.employees_tree.item(selection[0])
        employee_id = item['values'][0]
        
        try:
            conn = self.get_connection()
            if not conn:
                return
            
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM employees WHERE id = ?", (employee_id,))
            emp = cursor.fetchone()
            conn.close()
            
            if not emp:
                messagebox.showwarning("Внимание", "Сотрудник не найден")
                return
            
            # Создаем окно с детальной информацией
            dialog = tk.Toplevel(self.root)
            dialog.title(f"Информация о сотруднике: {emp[1]}")
            dialog.geometry("500x400")
            
            # Получаем названия колонок
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(employees)")
            columns = [col[1] for col in cursor.fetchall()]
            conn.close()
            
            # Отображаем информацию
            text = scrolledtext.ScrolledText(dialog, height=20)
            text.pack(fill='both', expand=True, padx=10, pady=10)
            
            info = ""
            for i, col in enumerate(columns):
                if i < len(emp):
                    if col == 'base_salary':
                        info += f"{col}: {emp[i]:,.2f} ₽\n"
                    elif col == 'is_active':
                        status = "Активен" if emp[i] else "Неактивен"
                        info += f"{col}: {status}\n"
                    else:
                        info += f"{col}: {emp[i]}\n"
            
            text.insert(1.0, info)
            text.config(state='disabled')
            
            # Кнопки
            button_frame = ttk.Frame(dialog)
            button_frame.pack(pady=10)
            
            ttk.Button(button_frame, text="Редактировать", 
                      command=lambda: [dialog.destroy(), self.show_edit_employee_dialog()]).pack(side='left', padx=10)
            ttk.Button(button_frame, text="Закрыть", 
                      command=dialog.destroy).pack(side='left', padx=10)
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить информацию:\n{e}")
    
    def delete_employee(self):
        """Удаление сотрудника"""
        selection = self.employees_tree.selection()
        if not selection:
            messagebox.showwarning("Внимание", "Выберите сотрудника для удаления")
            return
        
        item = self.employees_tree.item(selection[0])
        employee_id = item['values'][0]
        employee_name = item['values'][1]
        
        if messagebox.askyesno("Подтверждение", 
                              f"Вы уверены, что хотите удалить сотрудника:\n{employee_name}?"):
            try:
                conn = self.get_connection()
                if not conn:
                    return
                
                cursor = conn.cursor()
                cursor.execute("DELETE FROM employees WHERE id = ?", (employee_id,))
                conn.commit()
                conn.close()
                
                messagebox.showinfo("Успех", "Сотрудник успешно удален")
                self.load_employees()
                
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось удалить сотрудника:\n{e}")
    
    # ============================================================================
    # РАСЧЕТ ЗАРАБОТНОЙ ПЛАТЫ
    # ============================================================================
    
    def calculate_all_payroll(self):
        """Расчет зарплаты для всех сотрудников"""
        try:
            # Получаем параметры расчета
            month = self.month_var.get()
            total_days = int(self.working_days_var.get())
            worked_days = int(self.worked_days_var.get())
            kpi_score = self.bonus_kpi_var.get()
            overtime_hours = float(self.overtime_var.get())
            
            # Загружаем сотрудников
            conn = self.get_connection()
            if not conn:
                return
            
            cursor = conn.cursor()
            cursor.execute("SELECT id, full_name, position, base_salary FROM employees WHERE is_active = 1")
            employees = cursor.fetchall()
            conn.close()
            
            # Очищаем таблицу результатов
            for item in self.payroll_tree.get_children():
                self.payroll_tree.delete(item)
            
            total_income_sum = 0
            total_tax_sum = 0
            total_net_sum = 0
            
            # Рассчитываем для каждого сотрудника
            for emp in employees:
                employee = Employee(
                    id=emp[0],
                    full_name=emp[1],
                    position=emp[2],
                    base_salary=emp[3]
                )
                
                # Расчеты
                base_salary = self.calculator.calculate_base_salary(
                    employee, worked_days, total_days
                )
                
                bonus = self.calculator.calculate_bonus(employee, kpi_score)
                
                hourly_rate = employee.base_salary / 176  # 176 рабочих часов в месяце
                overtime_pay = self.calculator.calculate_overtime(overtime_hours, hourly_rate)
                
                total_income = self.calculator.calculate_total_income(base_salary, bonus, overtime_pay)
                tax_amount = self.tax_service.calculate_ndfl(total_income)
                net_salary = self.tax_service.calculate_net_salary(total_income, tax_amount)
                
                # Добавляем в таблицу
                self.payroll_tree.insert('', 'end', values=(
                    employee.full_name,
                    f"{base_salary:,.2f} ₽",
                    f"{bonus:,.2f} ₽",
                    f"{overtime_pay:,.2f} ₽",
                    f"{total_income:,.2f} ₽",
                    f"{tax_amount:,.2f} ₽",
                    f"{net_salary:,.2f} ₽"
                ))
                
                # Суммируем итоги
                total_income_sum += total_income
                total_tax_sum += tax_amount
                total_net_sum += net_salary
            
            # Обновляем итоги
            self.total_income_label.config(text=f"Общая сумма: {total_income_sum:,.2f} ₽")
            self.total_tax_label.config(text=f"Общий НДФЛ: {total_tax_sum:,.2f} ₽")
            self.total_net_label.config(text=f"К выплате: {total_net_sum:,.2f} ₽")
            
            messagebox.showinfo("Успех", f"Расчет завершен для {len(employees)} сотрудников")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось выполнить расчет:\n{e}")
    
    def calculate_selected_payroll(self):
        """Расчет зарплаты для выбранных сотрудников"""
        selection = self.employees_tree.selection()
        if not selection:
            messagebox.showwarning("Внимание", "Выберите сотрудников для расчета")
            return
        
        try:
            # Получаем параметры расчета
            month = self.month_var.get()
            total_days = int(self.working_days_var.get())
            worked_days = int(self.worked_days_var.get())
            kpi_score = self.bonus_kpi_var.get()
            overtime_hours = float(self.overtime_var.get())
            
            # Очищаем таблицу результатов
            for item in self.payroll_tree.get_children():
                self.payroll_tree.delete(item)
            
            total_income_sum = 0
            total_tax_sum = 0
            total_net_sum = 0
            
            conn = self.get_connection()
            if not conn:
                return
            
            cursor = conn.cursor()
            
            for item_id in selection:
                item = self.employees_tree.item(item_id)
                employee_id = item['values'][0]
                
                # Получаем данные сотрудника
                cursor.execute("SELECT id, full_name, position, base_salary FROM employees WHERE id = ?", (employee_id,))
                emp = cursor.fetchone()
                
                if emp:
                    employee = Employee(
                        id=emp[0],
                        full_name=emp[1],
                        position=emp[2],
                        base_salary=emp[3]
                    )
                    
                    # Расчеты
                    base_salary = self.calculator.calculate_base_salary(
                        employee, worked_days, total_days
                    )
                    
                    bonus = self.calculator.calculate_bonus(employee, kpi_score)
                    
                    hourly_rate = employee.base_salary / 176
                    overtime_pay = self.calculator.calculate_overtime(overtime_hours, hourly_rate)
                    
                    total_income = self.calculator.calculate_total_income(base_salary, bonus, overtime_pay)
                    tax_amount = self.tax_service.calculate_ndfl(total_income)
                    net_salary = self.tax_service.calculate_net_salary(total_income, tax_amount)
                    
                    # Добавляем в таблицу
                    self.payroll_tree.insert('', 'end', values=(
                        employee.full_name,
                        f"{base_salary:,.2f} ₽",
                        f"{bonus:,.2f} ₽",
                        f"{overtime_pay:,.2f} ₽",
                        f"{total_income:,.2f} ₽",
                        f"{tax_amount:,.2f} ₽",
                        f"{net_salary:,.2f} ₽"
                    ))
                    
                    # Суммируем итоги
                    total_income_sum += total_income
                    total_tax_sum += tax_amount
                    total_net_sum += net_salary
            
            conn.close()
            
            # Обновляем итоги
            self.total_income_label.config(text=f"Общая сумма: {total_income_sum:,.2f} ₽")
            self.total_tax_label.config(text=f"Общий НДФЛ: {total_tax_sum:,.2f} ₽")
            self.total_net_label.config(text=f"К выплате: {total_net_sum:,.2f} ₽")
            
            messagebox.showinfo("Успех", f"Расчет завершен для {len(selection)} сотрудников")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось выполнить расчет:\n{e}")
    
    def export_payroll_csv(self):
        """Экспорт результатов расчета в CSV"""
        try:
            # Собираем данные из таблицы
            data = []
            for item in self.payroll_tree.get_children():
                values = self.payroll_tree.item(item, 'values')
                if values:
                    data.append([v.replace(' ₽', '') for v in values])
            
            if not data:
                messagebox.showwarning("Внимание", "Нет данных для экспорта")
                return
            
            # Запрашиваем путь для сохранения
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV файлы", "*.csv"), ("Все файлы", "*.*")],
                initialfile=f"payroll_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            )
            
            if not filename:
                return
            
            # Сохраняем в CSV
            with open(filename, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file, delimiter=';')
                # Заголовки
                writer.writerow(['ФИО', 'Оклад', 'Премия', 'Сверхурочные', 
                               'Начислено', 'НДФЛ', 'К выплате'])
                # Данные
                for row in data:
                    writer.writerow(row)
            
            messagebox.showinfo("Успех", f"Данные успешно экспортированы в:\n{filename}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать данные:\n{e}")
    
    # ============================================================================
    # ОТЧЕТЫ
    # ============================================================================
    
    def generate_report(self, report_type):
        """Генерация отчета"""
        try:
            conn = self.get_connection()
            if not conn:
                return
            
            cursor = conn.cursor()
            report_content = ""
            
            if report_type == "payroll_report":
                # Расчетная ведомость
                cursor.execute("""
                    SELECT e.full_name, e.position, e.department, e.base_salary, 
                           e.bank_account, e.tax_id
                    FROM employees e 
                    WHERE e.is_active = 1
                    ORDER BY e.department, e.full_name
                """)
                
                employees = cursor.fetchall()
                
                report_content = "РАСЧЕТНАЯ ВЕДОМОСТЬ\n"
                report_content += "=" * 50 + "\n"
                report_content += f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
                report_content += f"Всего сотрудников: {len(employees)}\n"
                report_content += "=" * 50 + "\n\n"
                
                total_salary = 0
                for emp in employees:
                    report_content += f"ФИО: {emp[0]}\n"
                    report_content += f"Должность: {emp[1]}\n"
                    report_content += f"Отдел: {emp[2]}\n"
                    report_content += f"Оклад: {emp[3]:,.2f} ₽\n"
                    report_content += f"Банковский счет: {emp[4]}\n"
                    report_content += f"ИНН: {emp[5]}\n"
                    report_content += "-" * 30 + "\n"
                    total_salary += emp[3]
                
                report_content += f"\nОБЩИЙ ФОНД ОПЛАТЫ: {total_salary:,.2f} ₽\n"
                
            elif report_type == "employee_list":
                # Список сотрудников
                cursor.execute("""
                    SELECT e.full_name, e.position, e.department, e.base_salary, 
                           e.hire_date, e.email, e.phone
                    FROM employees e 
                    WHERE e.is_active = 1
                    ORDER BY e.department, e.full_name
                """)
                
                employees = cursor.fetchall()
                
                report_content = "СПИСОК СОТРУДНИКОВ\n"
                report_content += "=" * 50 + "\n"
                report_content += f"Дата формирования: {datetime.now().strftime('%d.%m.%Y')}\n"
                report_content += f"Всего сотрудников: {len(employees)}\n"
                report_content += "=" * 50 + "\n\n"
                
                report_content += f"{'ФИО':<30} {'Должность':<20} {'Отдел':<15} {'Оклад':>10} {'Дата приема':<12}\n"
                report_content += "-" * 87 + "\n"
                
                for emp in employees:
                    report_content += f"{emp[0]:<30} {emp[1]:<20} {emp[2]:<15} {emp[3]:>10,.0f} ₽ {emp[4]:<12}\n"
            
            elif report_type == "tax_report":
                # Налоговый отчет
                cursor.execute("""
                    SELECT e.full_name, e.tax_id, e.base_salary, 
                           e.base_salary * 0.13 as ndfl
                    FROM employees e 
                    WHERE e.is_active = 1
                    ORDER BY e.full_name
                """)
                
                employees = cursor.fetchall()
                
                report_content = "НАЛОГОВЫЙ ОТЧЕТ (НДФЛ)\n"
                report_content += "=" * 50 + "\n"
                report_content += f"Период: {datetime.now().strftime('%m.%Y')}\n"
                report_content += f"Ставка НДФЛ: 13%\n"
                report_content += "=" * 50 + "\n\n"
                
                report_content += f"{'ФИО':<30} {'ИНН':<15} {'Доход':>12} {'НДФЛ':>12}\n"
                report_content += "-" * 69 + "\n"
                
                total_income = 0
                total_tax = 0
                
                for emp in employees:
                    report_content += f"{emp[0]:<30} {emp[1]:<15} {emp[2]:>12,.2f} ₽ {emp[3]:>12,.2f} ₽\n"
                    total_income += emp[2]
                    total_tax += emp[3]
                
                report_content += "-" * 69 + "\n"
                report_content += f"{'ИТОГО:':<45} {total_income:>12,.2f} ₽ {total_tax:>12,.2f} ₽\n"
            
            elif report_type == "department_report":
                # Отчет по отделам
                cursor.execute("""
                    SELECT department, 
                           COUNT(*) as emp_count,
                           SUM(base_salary) as total_salary,
                           AVG(base_salary) as avg_salary
                    FROM employees 
                    WHERE is_active = 1
                    GROUP BY department
                    ORDER BY department
                """)
                
                departments = cursor.fetchall()
                
                report_content = "ОТЧЕТ ПО ОТДЕЛАМ\n"
                report_content += "=" * 50 + "\n"
                report_content += f"Дата формирования: {datetime.now().strftime('%d.%m.%Y')}\n"
                report_content += "=" * 50 + "\n\n"
                
                report_content += f"{'Отдел':<20} {'Сотр.':>6} {'ФОТ':>12} {'Средняя':>12}\n"
                report_content += "-" * 50 + "\n"
                
                total_employees = 0
                total_salary = 0
                
                for dept in departments:
                    report_content += f"{dept[0]:<20} {dept[1]:>6} {dept[2]:>12,.0f} ₽ {dept[3]:>12,.0f} ₽\n"
                    total_employees += dept[1]
                    total_salary += dept[2]
                
                report_content += "-" * 50 + "\n"
                report_content += f"{'ИТОГО:':<20} {total_employees:>6} {total_salary:>12,.0f} ₽\n"
            
            conn.close()
            
            # Отображаем отчет в текстовом поле
            self.report_text.delete(1.0, tk.END)
            self.report_text.insert(1.0, report_content)
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сгенерировать отчет:\n{e}")
    
    def save_report_csv(self):
        """Сохранение отчета как CSV"""
        try:
            report_content = self.report_text.get(1.0, tk.END).strip()
            if not report_content:
                messagebox.showwarning("Внимание", "Нет данных для сохранения")
                return
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV файлы", "*.csv"), ("Все файлы", "*.*")],
                initialfile=f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            )
            
            if not filename:
                return
            
            # Сохраняем как CSV
            with open(filename, 'w', newline='', encoding='utf-8') as file:
                # Преобразуем текстовый отчет в CSV
                lines = report_content.split('\n')
                for line in lines:
                    if line.strip():
                        # Заменяем множественные пробелы на разделитель
                        cells = [cell.strip() for cell in line.split('  ') if cell.strip()]
                        file.write(';'.join(cells) + '\n')
            
            messagebox.showinfo("Успех", f"Отчет сохранен в:\n{filename}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить отчет:\n{e}")
    
    def save_report_txt(self):
        """Сохранение отчета как TXT"""
        try:
            report_content = self.report_text.get(1.0, tk.END).strip()
            if not report_content:
                messagebox.showwarning("Внимание", "Нет данных для сохранения")
                return
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")],
                initialfile=f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            )
            
            if not filename:
                return
            
            # Сохраняем как TXT
            with open(filename, 'w', encoding='utf-8') as file:
                file.write(report_content)
            
            messagebox.showinfo("Успех", f"Отчет сохранен в:\n{filename}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить отчет:\n{e}")
    
    # ============================================================================
    # ИМПОРТ ИЗ 1С
    # ============================================================================
    
    def browse_import_file(self):
        """Выбор файла для импорта"""
        filetypes = [
            ("CSV файлы", "*.csv"),
            ("JSON файлы", "*.json"),
            ("Все файлы", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Выберите файл для импорта",
            filetypes=filetypes
        )
        
        if filename:
            self.file_path_var.set(filename)
            self.preview_import_file(filename)
    
    def preview_import_file(self, filename):
        """Предпросмотр файла импорта"""
        try:
            self.import_text.delete(1.0, tk.END)
            
            if filename.endswith('.csv'):
                with open(filename, 'r', encoding='utf-8') as file:
                    content = file.read()
                    # Показываем первые 1000 символов
                    preview = content[:1000]
                    self.import_text.insert(1.0, preview)
                    if len(content) > 1000:
                        self.import_text.insert(tk.END, "\n... (файл слишком большой для предпросмотра)")
            
            elif filename.endswith('.json'):
                with open(filename, 'r', encoding='utf-8') as file:
                    data = json.load(file)
                    preview = json.dumps(data, ensure_ascii=False, indent=2)
                    # Показываем первые 1000 символов
                    self.import_text.insert(1.0, preview[:1000])
                    if len(preview) > 1000:
                        self.import_text.insert(tk.END, "\n... (файл слишком большой для предпросмотра)")
            
            else:
                self.import_text.insert(1.0, "Неподдерживаемый формат файла")
                
        except Exception as e:
            self.import_text.insert(1.0, f"Ошибка чтения файла: {e}")
    
    def import_from_1c(self):
        """Импорт данных из 1С"""
        filename = self.file_path_var.get()
        if not filename or not os.path.exists(filename):
            messagebox.showwarning("Внимание", "Выберите файл для импорта")
            return
        
        try:
            imported_data = []
            
            if filename.endswith('.csv'):
                # Импорт из CSV
                with open(filename, 'r', encoding='utf-8') as file:
                    reader = csv.DictReader(file, delimiter=';')
                    for row in reader:
                        # Маппинг полей (адаптируйте под вашу структуру 1С)
                        imported_data.append({
                            'full_name': row.get('ФИО', ''),
                            'position': row.get('Должность', ''),
                            'base_salary': float(row.get('Оклад', 0)),
                            'department': row.get('Отдел', ''),
                            'tax_id': row.get('ИНН', ''),
                            'bank_account': row.get('БанковскийСчет', '')
                        })
            
            elif filename.endswith('.json'):
                # Импорт из JSON
                with open(filename, 'r', encoding='utf-8') as file:
                    data = json.load(file)
                    # Предполагаем, что данные в формате списка сотрудников
                    for item in data.get('employees', []):
                        imported_data.append({
                            'full_name': item.get('full_name', ''),
                            'position': item.get('position', ''),
                            'base_salary': float(item.get('base_salary', 0)),
                            'department': item.get('department', ''),
                            'tax_id': item.get('tax_id', ''),
                            'bank_account': item.get('bank_account', '')
                        })
            
            if not imported_data:
                messagebox.showwarning("Внимание", "В файле нет данных для импорта")
                return
            
            # Импортируем данные в базу
            conn = self.get_connection()
            if not conn:
                return
            
            cursor = conn.cursor()
            
            # Создаем таблицу если она не существует
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='employees'")
            if not cursor.fetchone():
                cursor.execute('''
                    CREATE TABLE employees (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        full_name TEXT NOT NULL,
                        position TEXT NOT NULL,
                        base_salary REAL NOT NULL,
                        department TEXT,
                        bank_account TEXT,
                        tax_id TEXT,
                        email TEXT,
                        phone TEXT,
                        hire_date TEXT,
                        is_active BOOLEAN DEFAULT TRUE,
                        created_at TEXT DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
            
            imported_count = 0
            updated_count = 0
            
            for emp_data in imported_data:
                # Проверяем, существует ли сотрудник
                cursor.execute("SELECT id FROM employees WHERE full_name = ?", (emp_data['full_name'],))
                existing = cursor.fetchone()
                
                if existing and self.update_existing_var.get():
                    # Обновляем существующего
                    cursor.execute('''
                        UPDATE employees 
                        SET position = ?, base_salary = ?, department = ?, 
                            bank_account = ?, tax_id = ?
                        WHERE id = ?
                    ''', (
                        emp_data['position'],
                        emp_data['base_salary'],
                        emp_data['department'],
                        emp_data['bank_account'],
                        emp_data['tax_id'],
                        existing[0]
                    ))
                    updated_count += 1
                elif self.create_missing_var.get():
                    # Добавляем нового
                    cursor.execute('''
                        INSERT INTO employees (full_name, position, base_salary, department, 
                                              bank_account, tax_id, hire_date)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        emp_data['full_name'],
                        emp_data['position'],
                        emp_data['base_salary'],
                        emp_data['department'],
                        emp_data['bank_account'],
                        emp_data['tax_id'],
                        datetime.now().strftime('%Y-%m-%d')
                    ))
                    imported_count += 1
            
            conn.commit()
            conn.close()
            
            # Обновляем список сотрудников
            self.load_employees()
            
            messagebox.showinfo("Успех", 
                              f"Импорт завершен:\n"
                              f"Импортировано новых: {imported_count}\n"
                              f"Обновлено существующих: {updated_count}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось импортировать данные:\n{e}")

# ============================================================================
# ЗАПУСК ПРИЛОЖЕНИЯ
# ============================================================================

def main():
    """Основная функция запуска приложения"""
    try:
        root = tk.Tk()
        
        # Устанавливаем стиль
        style = ttk.Style()
        style.theme_use('clam')
        
        # Настройка цветов
        root.configure(bg='white')
        
        # Создаем и запускаем приложение
        app = SalarySystemApp(root)
        
        # Центрируем окно
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'{width}x{height}+{x}+{y}')
        
        # Запуск главного цикла
        print("=" * 60)
        print("Система расчета заработной платы")
        print("Десктопная версия с редактированием")
        print("=" * 60)
        print("Приложение запущено в отдельном окне")
        print("Закройте окно приложения для выхода")
        print("=" * 60)
        
        root.mainloop()
        
    except Exception as e:
        print(f"Ошибка запуска приложения: {e}")
        input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()
        