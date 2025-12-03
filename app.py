import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# --- ГЛОБАЛЬНІ ЗМІННІ ---
# Шляхи до файлів тепер визначаються користувачем
FILE_PATH1 = None # Шлях до файлу-довідника (patients.xlsx)
df1 = None
df2 = None
changes = []
filtered_df = None  # Для зберігання фільтрованих даних на другій вкладці

# --- ФУНКЦІЇ ДЛЯ ОБРОБКИ ДАНИХ ТА GUI ---

def get_sheet_names(file_path):
    """Отримує список назв аркушів з Excel-файлу."""
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        sheets = wb.sheetnames
        wb.close()
        return sheets
    except Exception as e:
        messagebox.showerror("Помилка", f"Не вдалося прочитати аркуші файлу {file_path}: {e}")
        return []

def load_first_file_dialog():
    """Відкриває діалог вибору першого файлу та відображає вибір аркуша."""
    global FILE_PATH1
    
    selected_path = filedialog.askopenfilename(
        title="Вибрати файл-довідник (patients.xlsx)",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )
    
    if not selected_path:
        return # Користувач скасував вибір

    FILE_PATH1 = selected_path
    sheet_names = get_sheet_names(FILE_PATH1)
    
    if sheet_names:
        show_sheet_selection_dialog(sheet_names)
    else:
        FILE_PATH1 = None # Скидаємо шлях, якщо аркуші не завантажилися
        info_label.config(text="Помилка завантаження аркушів. Спробуйте інший файл.")


def show_sheet_selection_dialog(sheet_names):
    """Створює нове вікно для вибору аркуша."""
    
    sheet_window = tk.Toplevel(root)
    sheet_window.title("Вибір аркуша")
    sheet_window.geometry("300x150")
    sheet_window.transient(root) # Щоб вікно було поверх головного

    tk.Label(sheet_window, text="Оберіть аркуш для завантаження:", pady=10).pack()

    # Змінна для зберігання обраного значення
    selected_sheet = tk.StringVar(sheet_window)
    selected_sheet.set(sheet_names[0]) # Встановлюємо перший аркуш як дефолтний

    # Комбобокс для вибору аркуша
    sheet_chooser = ttk.Combobox(sheet_window, textvariable=selected_sheet, values=sheet_names, state="readonly")
    sheet_chooser.pack(pady=5)
    
    # Кнопка для підтвердження вибору
    def confirm_selection():
        sheet_name = selected_sheet.get()
        sheet_window.destroy()
        load_first_file_data(sheet_name) # Завантажуємо дані з обраного аркуша

    tk.Button(sheet_window, text="Завантажити", command=confirm_selection, 
              bg="#4CAF50", fg="white").pack(pady=10)


def load_first_file_data(sheet_name):
    """Завантажує дані з обраного аркуша першого файлу."""
    global df1
    
    try:
        df1 = pd.read_excel(FILE_PATH1, sheet_name=sheet_name)
        df1 = df1[['his_num', 'name']] # Вибираємо потрібні колонки
        
        info_label.config(text=f"Файл-довідник ('{FILE_PATH1.split('/')[-1]}', Аркуш: '{sheet_name}') завантажено.")
        # Активуємо кнопку завантаження другого файлу
        load_button.config(state=tk.NORMAL)
    except Exception as e:
        df1 = None
        messagebox.showerror("Помилка завантаження", f"Помилка при завантаженні даних з {FILE_PATH1} ({sheet_name}): {e}")
        load_button.config(state=tk.DISABLED)


def load_and_process_files():
    """Викликається після вибору другого файлу. Обробляє та відображає дані."""
    global df2, changes
    
    if df1 is None:
        messagebox.showwarning("Попередження", "Спочатку успішно завантажте файл-довідник.")
        return
        
    file_path2 = filedialog.askopenfilename(
        title="Вибрати файл для оновлення (hospitalizations_11.xlsx)",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )
    
    if not file_path2:
        return 
        
    try:
        df2 = pd.read_excel(file_path2)
    except Exception as e:
        messagebox.showerror("Помилка завантаження", f"Помилка при завантаженні {file_path2}: {e}")
        return

    # 2. Обробка даних (як у оригінальному коді)
    # Перейменування колонок
    df2_renamed = df2.rename(columns={
        'Номер паперової історії хвороби': 'his_num',
        'ПІБ пацієнта': 'name'
    })

    # Об'єднання
    merged_df = df2_renamed.merge(
        df1[['his_num', 'name']], 
        on='his_num', 
        how='left', 
        suffixes=('_old', '_new')
    )

    # Пошук змін
    changes = []
    for idx, row in merged_df.iterrows():
        if pd.notna(row['name_new']) and row['name_old'] != row['name_new']:
            changes.append({
                'his_num': row['his_num'],
                'old_name': row['name_old'],
                'new_name': row['name_new'],
                'index': idx 
            })
            
    # 3. Оновлення GUI
    update_treeview()
    info_label.config(text=f"Файл '{file_path2.split('/')[-1]}' завантажено. Знайдено змін: {len(changes)}")
    save_button.config(state=tk.NORMAL)


def update_treeview():
    """Очищує та заповнює таблицю (Treeview) новими даними."""
    for item in tree.get_children():
        tree.delete(item)
        
    for i, change in enumerate(changes, 1):
        tree.insert('', tk.END, values=(
            i,
            change['his_num'],
            change['old_name'],
            change['new_name']
        ))

def load_hospitalization_file():
    """Завантажує файл госпіталізацій та відображає відфільтровані дані."""
    global filtered_df
    
    file_path = filedialog.askopenfilename(
        title="Вибрати файл госпіталізацій",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )
    
    if not file_path:
        return
    
    try:
        df = pd.read_excel(file_path)
        
        # Вибираємо потрібні колонки
        selected_columns = df[[
            'Номер паперової історії хвороби',
            'Дата госпіталізації',
            'Медпрацівник. відповідальний за епізод',
            'ПІБ пацієнта',
            'Дата та час виписки',
            'Помилка оплати НСЗУ'
        ]]
        
        # Фільтруємо дані - залишаємо тільки рядки, де 'Помилка оплати НСЗУ' не пуста
        filtered_df = selected_columns[selected_columns['Помилка оплати НСЗУ'].notnull()]
        
        # Оновлюємо таблицю на вкладці 2
        update_filtered_treeview()
        
        status_label.config(text=f"Завантажено! Знайдено {len(filtered_df)} записів з помилками оплати", fg="green")
        
    except KeyError as e:
        messagebox.showerror("Помилка", f"Файл не містить необхідної колонки: {e}")
    except Exception as e:
        messagebox.showerror("Помилка", f"Помилка при завантаженні файлу: {e}")

def update_filtered_treeview():
    """Оновлює таблицю з фільтрованими даними."""
    if filtered_df is None:
        return
    
    # Очищаємо таблицю
    for item in tree_filtered.get_children():
        tree_filtered.delete(item)
    
    # Заповнюємо таблицю
    for idx, row in filtered_df.iterrows():
        tree_filtered.insert('', tk.END, values=(
            idx + 1,
            row['Номер паперової історії хвороби'],
            row['Дата госпіталізації'],
            row['Медпрацівник. відповідальний за епізод'],
            row['ПІБ пацієнта'],
            row['Дата та час виписки'],
            row['Помилка оплати НСЗУ']
        ))

def save_changes():
    """Зберігає зміни в новий Excel-файл з виділенням."""
    if df2 is None or len(changes) == 0:
        messagebox.showinfo("Інформація", "Немає змін для збереження або не завантажено файл для оновлення.")
        return
    
    # ... (Логіка збереження і форматування залишається без змін) ...
    df3 = df2.copy()
    name_changes = {change['his_num']: change['new_name'] for change in changes}
    changed_indices = []
    
    for idx, row in df3.iterrows():
        his_num = row.get('Номер паперової історії хвороби')
        if his_num in name_changes:
            df3.at[idx, 'ПІБ пацієнта'] = name_changes[his_num]
            changed_indices.append(idx) 
    
    output_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile='hospitalizations_11_updated.xlsx',
        title="Зберегти оновлений файл як...",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )

    if not output_file:
        return 
        
    df3.to_excel(output_file, index=False)
    
    try:
        wb = openpyxl.load_workbook(output_file)
        ws = wb.active
    except Exception as e:
        messagebox.showerror("Помилка форматування", f"Не вдалося відкрити файл для форматування: {e}")
        return

    name_col_idx = None
    for idx, col in enumerate(ws.iter_cols(1, ws.max_column, 1, 1), start=1):
        if col[0].value == 'ПІБ пацієнта':
            name_col_idx = idx
            break
    
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    
    if name_col_idx:
        for idx in changed_indices:
            cell = ws.cell(row=idx + 2, column=name_col_idx) 
            cell.fill = yellow_fill
    
    wb.save(output_file)
    
    messagebox.showinfo("Успіх", 
                        f"Файл '{output_file.split('/')[-1]}' успішно збережено!\n"
                        f"Змінено записів: {len(changed_indices)}")


# --- СТВОРЕННЯ GUI ---
root = tk.Tk()
root.title("Перегляд та оновлення ПІБ пацієнтів")
root.geometry("900x600")

# --- СТВОРЕННЯ ВКЛАДОК (NOTEBOOK) ---
notebook = ttk.Notebook(root)
notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

# ===== ВКЛАДКА 1: Основна функціональність =====
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="Основна")

title_label = tk.Label(tab1, text="Знайдені розбіжності в ПІБ пацієнтів", 
                        font=("Arial", 14, "bold"), pady=10)
title_label.pack()

info_label = tk.Label(tab1, text="1. Виберіть файл-довідник. 2. Виберіть файл для оновлення.", 
                        font=("Arial", 10), pady=5)
info_label.pack()

# --- КНОПКИ УПРАВЛІННЯ ---
button_frame = tk.Frame(tab1)
button_frame.pack(pady=10)

# 1. Кнопка вибору основного файлу
load_file1_button = tk.Button(button_frame, text="Вибрати файл-довідник", 
                              command=load_first_file_dialog, bg="#FF9800", fg="white",
                              font=("Arial", 10, "bold"), padx=10, pady=5)
load_file1_button.pack(side=tk.LEFT, padx=5)

# 2. Кнопка для завантаження другого файлу (заблокована до вибору першого)
load_button = tk.Button(button_frame, text="Вибрати файл H24", 
                        command=load_and_process_files, bg="#2196F3", fg="white",
                        font=("Arial", 10, "bold"), padx=10, pady=5, state=tk.DISABLED)
load_button.pack(side=tk.LEFT, padx=5)

# 3. Кнопка для збереження змін (заблокована до знаходження змін)
save_button = tk.Button(button_frame, text="Зберегти зміни", 
                        command=save_changes, bg="#4CAF50", fg="white",
                        font=("Arial", 10, "bold"), padx=10, pady=5, state=tk.DISABLED)
save_button.pack(side=tk.LEFT, padx=5)

exit_button = tk.Button(button_frame, text="Вихід", 
                        command=root.quit, bg="#f44336", fg="white",
                        font=("Arial", 10, "bold"), padx=10, pady=5)
exit_button.pack(side=tk.LEFT, padx=5)

# --- ТАБЛИЦЯ (TREEVIEW) ---
frame = ttk.Frame(tab1)
frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

scrollbar = ttk.Scrollbar(frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

columns = ("№", "Номер історії", "Старе ПІБ", "Нове ПІБ")
tree = ttk.Treeview(frame, columns=columns, show='headings', yscrollcommand=scrollbar.set)
scrollbar.config(command=tree.yview)

tree.heading("№", text="№")
tree.heading("Номер історії", text="Номер історії")
tree.heading("Старе ПІБ", text="Старе ПІБ")
tree.heading("Нове ПІБ", text="Нове ПІБ")

tree.column("№", width=50, anchor='center')
tree.column("Номер історії", width=120, anchor='center')
tree.column("Старе ПІБ", width=300)
tree.column("Нове ПІБ", width=300)

tree.pack(fill=tk.BOTH, expand=True)

# ===== ВКЛАДКА 2: Фільтровані дані по помилкам оплати =====
tab2 = ttk.Frame(notebook)
notebook.add(tab2, text="Помилки оплати")

# Заголовок та кнопка завантаження файлу
header_frame = tk.Frame(tab2, bg="lightblue", padx=10, pady=10)
header_frame.pack(fill=tk.X)

title_label2 = tk.Label(header_frame, text="Записи з помилками оплати НСЗУ", 
                        font=("Arial", 12, "bold"), bg="lightblue")
title_label2.pack(side=tk.LEFT)

load_file_button2 = tk.Button(header_frame, text="Завантажити файл", 
                              command=load_hospitalization_file, bg="#4CAF50", fg="white",
                              font=("Arial", 10, "bold"), padx=15, pady=5)
load_file_button2.pack(side=tk.RIGHT, padx=5)

# Статус
status_label = tk.Label(tab2, text="Файл не завантажено", 
                        font=("Arial", 9), fg="#666")
status_label.pack(pady=5)

# Таблиця з фільтрованими даними
frame_filtered = ttk.Frame(tab2)
frame_filtered.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

scrollbar_filtered = ttk.Scrollbar(frame_filtered)
scrollbar_filtered.pack(side=tk.RIGHT, fill=tk.Y)

columns_filtered = ("№", "Номер історії", "Дата госпіталізації", "Медпрацівник", "ПІБ пацієнта", "Дата виписки", "Помилка оплати")
tree_filtered = ttk.Treeview(frame_filtered, columns=columns_filtered, show='headings', yscrollcommand=scrollbar_filtered.set)
scrollbar_filtered.config(command=tree_filtered.yview)

tree_filtered.heading("№", text="№")
tree_filtered.heading("Номер історії", text="Номер історії")
tree_filtered.heading("Дата госпіталізації", text="Дата госпіталізації")
tree_filtered.heading("Медпрацівник", text="Медпрацівник")
tree_filtered.heading("ПІБ пацієнта", text="ПІБ пацієнта")
tree_filtered.heading("Дата виписки", text="Дата виписки")
tree_filtered.heading("Помилка оплати", text="Помилка оплати")

tree_filtered.column("№", width=40, anchor='center')
tree_filtered.column("Номер історії", width=80, anchor='center')
tree_filtered.column("Дата госпіталізації", width=100, anchor='center')
tree_filtered.column("Медпрацівник", width=150)
tree_filtered.column("ПІБ пацієнта", width=150)
tree_filtered.column("Дата виписки", width=100, anchor='center')
tree_filtered.column("Помилка оплати", width=200)

tree_filtered.pack(fill=tk.BOTH, expand=True)

# --- ЗАПУСК GUI ---
root.mainloop()