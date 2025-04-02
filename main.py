import sys
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from tkinter import filedialog
from modules.act_processor import ActProcessor
from modules.file_manager import FileManager
from config import DEFAULT_TEMPLATE, DEFAULT_SOURCE, DEFAULT_REGISTER, OUTPUT_DIR

class AktGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.processor = ActProcessor()
        self.file_manager = FileManager()
        self.setup_ui()
        
        # Переменные состояния
        self.valid_rows = []
        self.current_output_dir = OUTPUT_DIR

    def setup_ui(self):
        """Настройка графического интерфейса"""
        self.root.title("Генератор актов скрытых работ")
        self.root.geometry("900x700")
        
        # Основные фреймы
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Панель управления
        control_frame = ttk.LabelFrame(main_frame, text="Управление", padding="10")
        control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(control_frame, 
                 text="Загрузить реестр", 
                 command=self.load_register).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame,
                 text="Сгенерировать акты",
                 command=self.generate_akts).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame,
                 text="Выбрать папку для сохранения",
                 command=self.select_output_dir).pack(side=tk.LEFT, padx=5)
        
        # Лог выполнения
        self.log_frame = ttk.LabelFrame(main_frame, text="Лог выполнения", padding="10")
        self.log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(self.log_frame, wrap=tk.WORD, state='disabled')
        scrollbar = ttk.Scrollbar(self.log_frame, command=self.log_text.yview)
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

    def load_register(self):
        """Загрузка файла реестра"""
        try:
            self.log_message("\nЗагрузка реестра...")
            self.valid_rows = self.processor.process_register(DEFAULT_REGISTER)
            
            if not self.valid_rows:
                messagebox.showwarning("Внимание", "Реестр не содержит данных для обработки")
                return
            
            self.log_message(f"Успешно загружено {len(self.valid_rows)} актов для обработки")
            self.status_var.set(f"Готово к генерации: {len(self.valid_rows)} актов")
            
        except Exception as e:
            self.log_message(f"Ошибка загрузки реестра: {str(e)}")
            messagebox.showerror("Ошибка", f"Не удалось загрузить реестр:\n{str(e)}")

    def select_output_dir(self):
        """Выбор папки для сохранения результатов"""
        dir_path = filedialog.askdirectory(initialdir=self.current_output_dir)
        if dir_path:
            self.current_output_dir = Path(dir_path)
            self.log_message(f"Папка для сохранения изменена на: {self.current_output_dir}")

    def generate_akts(self):
        """Основная функция генерации актов"""
        if not self.valid_rows:
            messagebox.showwarning("Внимание", "Сначала загрузите реестр актов")
            return
        
        success_count = 0
        self.log_message("\nНачало генерации актов...")
        
        for i, row in enumerate(self.valid_rows, 1):
            akt_num = f"{self.file_manager.safe_get(row, 0, '')}{self.file_manager.safe_get(row, 1, '')}"
            self.log_message(f"\nОбработка акта {i}/{len(self.valid_rows)}: {akt_num}")
            
            result = self.processor.generate_akt(
                row=row,
                template_path=DEFAULT_TEMPLATE,
                source_path=DEFAULT_SOURCE,
                output_dir=self.current_output_dir
            )
            
            if result['status'] == 'success':
                success_count += 1
                self.log_message(f"Успешно создан: {result['file']}")
            else:
                self.log_message(f"Ошибка: {result['error']}")
        
        self.log_message("\n" + "="*50)
        self.log_message(f"ГЕНЕРАЦИЯ ЗАВЕРШЕНА\nУспешно создано: {success_count}/{len(self.valid_rows)}")
        self.status_var.set(f"Готово. Успешно создано {success_count} актов")
        
        messagebox.showinfo("Завершено", 
                          f"Обработка завершена.\nУспешно создано актов: {success_count}/{len(self.valid_rows)}")

def main():
    try:
        # Создаем папку для результатов если ее нет
        OUTPUT_DIR.mkdir(exist_ok=True)
        
        # Создаем и запускаем GUI
        root = tk.Tk()
        app = AktGeneratorGUI(root)
        root.mainloop()
        
    except Exception as e:
        print(f"Критическая ошибка: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()