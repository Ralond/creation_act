import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from modules.act_processor import ActProcessor
from modules.file_manager import FileManager
from config import DEFAULT_REGISTER, OUTPUT_DIR

class AktGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.processor = ActProcessor()
        self.file_manager = FileManager()
        self.setup_ui()
        self.current_akt = None
        
    def setup_ui(self):
        """Настройка интерфейса"""
        self.root.title("Генератор актов")
        self.root.geometry("900x700")
        
        # Основные фреймы
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Панель управления
        control_frame = ttk.LabelFrame(main_frame, text="Управление", padding="10")
        control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(control_frame, text="Выбрать реестр", command=self.select_register).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Сгенерировать акты", command=self.generate_akts).pack(side=tk.LEFT, padx=5)
        
        # Панель сертификатов
        cert_frame = ttk.LabelFrame(main_frame, text="Добавление сертификатов", padding="10")
        cert_frame.pack(fill=tk.X, pady=5)
        
        self.akt_combobox = ttk.Combobox(cert_frame, state='readonly')
        self.akt_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        ttk.Button(cert_frame, text="Добавить сертификат", command=self.add_certificate).pack(side=tk.LEFT, padx=5)
        
        # Лог действий
        log_frame = ttk.LabelFrame(main_frame, text="Лог выполнения", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, state='disabled')
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Статус бар
        self.status_var = tk.StringVar(value="Готов к работе")
        ttk.Label(main_frame, textvariable=self.status_var).pack(side=tk.BOTTOM, fill=tk.X)
    
    def log_message(self, message):
        """Вывод сообщения в лог"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)
        self.root.update()
    
    def select_register(self):
        """Выбор файла реестра"""
        filepath = filedialog.askopenfilename(
            initialfile=DEFAULT_REGISTER,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if filepath:
            self.log_message(f"Выбран реестр: {filepath}")
            self.register_path = Path(filepath)
    
    def generate_akts(self):
        """Генерация всех актов"""
        try:
            if not hasattr(self, 'register_path'):
                messagebox.showwarning("Внимание", "Сначала выберите файл реестра!")
                return
                
            self.log_message("\nНачало обработки реестра...")
            rows = self.processor.process_register(self.register_path)
            
            if not rows:
                messagebox.showinfo("Информация", "Нет данных для обработки в реестре")
                return
                
            self.log_message(f"Найдено актов для обработки: {len(rows)}")
            
            # Обновляем комбобокс
            akt_numbers = [f"{self.file_manager.safe_get(row, 0, '')}{self.file_manager.safe_get(row, 1, '')}" for row in rows]
            self.akt_combobox['values'] = akt_numbers
            
            self.log_message("\nГенерация актов завершена успешно!")
            messagebox.showinfo("Успех", f"Сгенерировано {len(rows)} актов")
            
        except Exception as e:
            self.log_message(f"\nОшибка: {str(e)}")
            messagebox.showerror("Ошибка", str(e))
    
    def add_certificate(self):
        """Добавление сертификата к акту"""
        akt_num = self.akt_combobox.get()
        if not akt_num:
            messagebox.showwarning("Внимание", "Сначала выберите акт!")
            return
            
        filepath = filedialog.askopenfilename(
            title=f"Выберите сертификат для акта {akt_num}",
            filetypes=[("PDF/Изображения", "*.pdf *.jpg *.jpeg *.png")]
        )
        
        if filepath:
            filepath = Path(filepath)
            if self.file_manager.validate_certificate(filepath):
                # Здесь логика добавления сертификата
                self.log_message(f"Добавлен сертификат: {filepath.name} к акту {akt_num}")
            else:
                messagebox.showerror("Ошибка", "Неподдерживаемый формат файла!")