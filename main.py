import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from modules.act_processor import ActProcessor
from modules.file_manager import FileManager
from config import OUTPUT_DIR

class AktGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.processor = ActProcessor()
        self.file_manager = FileManager()
        self.current_output_dir = OUTPUT_DIR
        self.setup_ui()

    def setup_ui(self):
        """Настройка графического интерфейса"""
        self.root.title("Генератор актов скрытых работ (АОСР)")
        self.root.geometry("1000x800")
        
        # Основные фреймы
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Панель управления
        control_frame = ttk.LabelFrame(main_frame, text="Управление", padding="10")
        control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(control_frame, 
                 text="Загрузить реестр АОСР", 
                 command=self.load_register).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame,
                 text="Выбрать шаблон акта",
                 command=self.select_template).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame,
                 text="Сгенерировать акты",
                 command=self.generate_akts).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame,
                 text="Выбрать папку для сохранения",
                 command=self.select_output_dir).pack(side=tk.LEFT, padx=5)
        
        # Информация о выбранных файлах
        info_frame = ttk.LabelFrame(main_frame, text="Информация", padding="10")
        info_frame.pack(fill=tk.X, pady=5)
        
        self.register_label = ttk.Label(info_frame, text="Реестр АОСР: не выбран")
        self.register_label.pack(anchor=tk.W)
        
        self.template_label = ttk.Label(info_frame, text="Шаблон акта: не выбран")
        self.template_label.pack(anchor=tk.W)
        
        self.output_label = ttk.Label(info_frame, text=f"Папка для сохранения: {self.current_output_dir}")
        self.output_label.pack(anchor=tk.W)
        
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
        filepath = filedialog.askopenfilename(
            title="Выберите файл реестра АОСР",
            filetypes=[("Excel files", "*.xlsm *.xlsx")]
        )
        
        if filepath:
            self.register_path = Path(filepath)
            self.register_label.config(text=f"Реестр АОСР: {filepath}")
            self.log_message(f"Выбран реестр: {filepath}")

    def select_template(self):
        """Выбор шаблона акта"""
        filepath = filedialog.askopenfilename(
            title="Выберите шаблон акта",
            filetypes=[("Excel files", "*.xls *.xlsx")]
        )
        
        if filepath:
            self.template_path = Path(filepath)
            self.template_label.config(text=f"Шаблон акта: {filepath}")
            self.log_message(f"Выбран шаблон: {filepath}")

    def select_output_dir(self):
        """Выбор папки для сохранения результатов"""
        dir_path = filedialog.askdirectory(initialdir=self.current_output_dir)
        if dir_path:
            self.current_output_dir = Path(dir_path)
            self.output_label.config(text=f"Папка для сохранения: {self.current_output_dir}")
            self.log_message(f"Папка для сохранения изменена на: {self.current_output_dir}")

    def generate_akts(self):
        """Основная функция генерации актов"""
        if not hasattr(self, 'register_path'):
            messagebox.showwarning("Внимание", "Сначала выберите файл реестра АОСР!")
            return
        
        if not hasattr(self, 'template_path'):
            messagebox.showwarning("Внимание", "Сначала выберите шаблон акта!")
            return
    
        try:
        # Запрашиваем место сохранения итогового файла
            output_file = filedialog.asksaveasfilename(
                title="Сохранить книгу актов",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
        
            if not output_file:
                return
            
            output_path = Path(output_file)
        
            self.log_message("\nНачало обработки реестра...")
            rows = self.processor.process_register(self.register_path)
        
            if not rows:
                messagebox.showinfo("Информация", "Реестр не содержит данных для обработки")
                return
        
            self.log_message(f"Найдено актов для обработки: {len(rows)}")
        
        # Генерация всех актов в одной книге
            result = self.processor.generate_all_akts(
                rows=rows,
                template_path=self.template_path,
                output_path=output_path
            )
        
            if result['status'] == 'success':
                self.log_message("\n" + "="*50)
                self.log_message(f"ГЕНЕРАЦИЯ ЗАВЕРШЕНА\nУспешно создано: {result['success']}/{result['total']}")
                self.log_message(f"Результат сохранен в: {output_path}")
                self.status_var.set(f"Готово. Успешно создано {result['success']} актов")
            
                messagebox.showinfo("Завершено", 
                              f"Обработка завершена.\nУспешно создано актов: {result['success']}/{result['total']}\n\nФайл сохранен:\n{output_path}")
            else:
                self.log_message(f"\nОшибка: {result['error']}")
                messagebox.showerror("Ошибка", f"Произошла ошибка:\n{result['error']}")
            
        except Exception as e:
            self.log_message(f"\nОшибка: {str(e)}")
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")
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