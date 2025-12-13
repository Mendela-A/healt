import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class NZSUFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Фільтр помилок оплати НСЗУ")
        self.root.geometry("1400x700")
        
        # Центрування вікна
        self.center_window()
        
        self.file_path = None
        self.df = None
        self.filtered_data = []
        
        self.create_widgets()
    
    def center_window(self):
        """Центрування вікна на екрані"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        # Фрейм для завантаження файлу
        file_frame = ttk.LabelFrame(self.root, text="Файл даних", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(file_frame, text="Завантажити файл", 
                   command=self.load_file).pack(side="left", padx=5)
        self.file_label = ttk.Label(file_frame, text="Файл не вибрано")
        self.file_label.pack(side="left", padx=5)
        
        # Фрейм для вибору листа
        sheet_frame = ttk.Frame(file_frame)
        sheet_frame.pack(side="left", padx=20)
        ttk.Label(sheet_frame, text="Лист:").pack(side="left")
        self.sheet_combo = ttk.Combobox(sheet_frame, width=20, state="disabled")
        self.sheet_combo.pack(side="left", padx=5)
        
        # Кнопка фільтрації
        ttk.Button(self.root, text="Фільтрувати дані", 
                   command=self.filter_data).pack(pady=10)
        
        # Фрейм для суми та кнопки збереження (створюємо ПЕРЕД таблицею)
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(side="bottom", fill="x", padx=10, pady=10)
        
        self.sum_label = ttk.Label(bottom_frame, text="Загальна сума: 0.00 грн", 
                                    font=("Arial", 12, "bold"))
        self.sum_label.pack(side="left", padx=10)
        
        ttk.Button(bottom_frame, text="Зберегти результати", 
                   command=self.save_results).pack(side="right", padx=10)
        
        # Фрейм для таблиці результатів (створюємо ПІСЛЯ bottom_frame)
        result_frame = ttk.LabelFrame(self.root, text="Результати", padding=10)
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Створення Treeview (таблиця)
        columns = ("Номер історії", "Медпрацівник", "Оплата НСЗУ (грн)", "Помилка оплати НСЗУ")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=20)
        
        # Налаштування стилю для багаторядкових комірок
        style = ttk.Style()
        style.configure("Treeview", rowheight=40)
        
        # Налаштування тегів для кольорового виділення
        self.tree.tag_configure('death', background='#ffcccc', foreground='#cc0000')
        
        # Налаштування колонок
        self.tree.heading("Номер історії", text="Номер історії")
        self.tree.heading("Медпрацівник", text="Медпрацівник")
        self.tree.heading("Оплата НСЗУ (грн)", text="Оплата НСЗУ (грн)")
        self.tree.heading("Помилка оплати НСЗУ", text="Помилка оплати НСЗУ")
        
        self.tree.column("Номер історії", width=100, anchor="center")
        self.tree.column("Медпрацівник", width=200)
        self.tree.column("Оплата НСЗУ (грн)", width=120, anchor="center")
        self.tree.column("Помилка оплати НСЗУ", width=600)
        
        # Scrollbar для таблиці
        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="Виберіть файл",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=file_path.split('/')[-1])
            
            # Завантаження списку листів
            xl_file = pd.ExcelFile(file_path)
            self.sheet_combo['values'] = xl_file.sheet_names
            self.sheet_combo.current(0)
            self.sheet_combo['state'] = "readonly"
    
    def normalize_his_num(self, num):
        """Видаляє провідні нулі з номера історії"""
        try:
            return str(int(float(str(num))))
        except:
            return str(num).strip()
    
    def filter_data(self):
        if not self.file_path:
            messagebox.showwarning("Увага", "Завантажте файл")
            return
        
        try:
            # Очищення таблиці
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Завантаження файлу
            sheet_name = self.sheet_combo.get()
            self.df = pd.read_excel(self.file_path, sheet_name=sheet_name)
            
            # Перевірка колонок
            required_cols = [
                'Помилка оплати НСЗУ',
                'Номер паперової історії хвороби',
                'Медпрацівник. відповідальний за епізод',
                'Оплата НСЗУ (грн)',
                'Результат лікування'
            ]
            
            missing_cols = [col for col in required_cols if col not in self.df.columns]
            if missing_cols:
                messagebox.showerror("Помилка", f"Файл не містить колонок: {', '.join(missing_cols)}")
                return
            
            # Фільтрація даних (тільки рядки з помилками)
            filtered_df = self.df[self.df['Помилка оплати НСЗУ'].notna() & 
                                  (self.df['Помилка оплати НСЗУ'] != '')]
            
            self.filtered_data = []
            total_sum = 0
            
            for idx, row in filtered_df.iterrows():
                his_num = self.normalize_his_num(row['Номер паперової історії хвороби'])
                medworker = str(row['Медпрацівник. відповідальний за епізод'])
                payment = row['Оплата НСЗУ (грн)']
                error = str(row['Помилка оплати НСЗУ'])
                result = str(row['Результат лікування']) if pd.notna(row['Результат лікування']) else ''
                
                # Обробка суми
                try:
                    payment_value = float(payment) if pd.notna(payment) else 0
                except:
                    payment_value = 0
                
                total_sum += payment_value
                
                self.filtered_data.append({
                    'Номер історії': his_num,
                    'Медпрацівник': medworker,
                    'Оплата НСЗУ (грн)': payment_value,
                    'Помилка оплати НСЗУ': error
                })
                
                # Перевірка на смерть пацієнта для червоного виділення
                is_death = 'помер' in result.lower()
                
                # Додавання рядка в таблицю з тегом для червоного кольору
                if is_death:
                    self.tree.insert("", "end", values=(
                        his_num, 
                        medworker, 
                        f"{payment_value:.2f}",
                        error
                    ), tags=('death',))
                else:
                    self.tree.insert("", "end", values=(
                        his_num, 
                        medworker, 
                        f"{payment_value:.2f}",
                        error
                    ))
            
            # Оновлення суми
            self.sum_label.config(text=f"Загальна сума: {total_sum:.2f} грн")
            
            if self.filtered_data:
                messagebox.showinfo("Результат", f"Знайдено {len(self.filtered_data)} записів з помилками")
            else:
                messagebox.showinfo("Результат", "✓ Записів з помилками не знайдено!")
            
        except Exception as e:
            messagebox.showerror("Помилка", f"Виникла помилка: {str(e)}")
    
    def save_results(self):
        if not self.filtered_data:
            messagebox.showwarning("Увага", "Немає результатів для збереження")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if file_path:
            df_result = pd.DataFrame(self.filtered_data)
            
            # Додавання рядка з сумою
            total_sum = df_result['Оплата НСЗУ (грн)'].sum()
            sum_row = pd.DataFrame([{
                'Номер історії': '',
                'Медпрацівник': '',
                'Оплата НСЗУ (грн)': total_sum,
                'Помилка оплати НСЗУ': 'ЗАГАЛЬНА СУМА'
            }])
            
            df_result = pd.concat([df_result, sum_row], ignore_index=True)
            df_result.to_excel(file_path, index=False)
            messagebox.showinfo("Успіх", f"Результати збережено у файл:\n{file_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = NZSUFilterApp(root)
    root.mainloop()