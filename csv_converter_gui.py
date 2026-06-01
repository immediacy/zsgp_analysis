import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os


def converter(a_path, separator_sign):
    """Конвертирует CSV файл с указанным разделителем"""
    try:
        clean_path = a_path.replace('"', '').strip()

        # Проверяем существование файла
        if not os.path.exists(clean_path):
            return 1

        # Читаем CSV с разделителем ","
        df = pd.read_csv(clean_path, sep=',')

        # Формируем имя выходного файла
        base_name = clean_path.replace('.csv', '')
        output_path = f"{base_name}_separated.csv"

        # Сохраняем с новым разделителем
        df.to_csv(output_path, index=False, sep=separator_sign)

        return 0, output_path
    except FileNotFoundError:
        return 1, None
    except Exception as e:
        return 2, str(e)


class CSVConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер CSV файлов")
        self.root.geometry("600x400")
        self.root.resizable(True, True)

        # Настройка стилей
        style = ttk.Style()
        style.theme_use('clam')

        # Главный фрейм с отступами
        main_frame = ttk.Frame(root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Настройка растягивания колонок
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Заголовок
        title_label = ttk.Label(main_frame, text="Конвертер CSV файлов",
                                font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # 1. Путь к файлу
        ttk.Label(main_frame, text="Путь к CSV файлу:", font=('Arial', 10)).grid(
            row=1, column=0, sticky=tk.W, pady=(0, 5))

        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(main_frame, textvariable=self.path_var,
                                    font=('Arial', 10), width=50)
        self.path_entry.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        # Кнопка выбора файла
        self.browse_btn = ttk.Button(main_frame, text="Обзор...",
                                     command=self.browse_file)
        self.browse_btn.grid(row=2, column=2, padx=(10, 0), pady=(0, 10))

        # 2. Выпадающий список с разделителями
        ttk.Label(main_frame, text="Выберите целевой разделитель:", font=('Arial', 10)).grid(
            row=3, column=0, sticky=tk.W, pady=(10, 5))

        self.separator_map = {
            "точка с запятой": ";",
            "запятая": ",",
            "знак табуляции": "\t"
        }

        self.sep_var = tk.StringVar(value="точка с запятой")
        self.sep_combo = ttk.Combobox(main_frame, textvariable=self.sep_var,
                                      values=list(self.separator_map.keys()),
                                      state="readonly", width=30, font=('Arial', 10))
        self.sep_combo.grid(row=4, column=0, sticky=tk.W, pady=(0, 20))

        # Информационная метка
        info_text = "Примечание: исходный файл должен иметь разделитель 'запятая' (',')\n"
        info_text += "Результат сохранится в файл с суффиксом '_separated.csv'"
        info_label = ttk.Label(main_frame, text=info_text, font=('Arial', 9),
                               foreground='gray')
        info_label.grid(row=5, column=0, columnspan=3, pady=(0, 20))

        # Кнопка конвертации
        self.convert_btn = ttk.Button(main_frame, text="Конвертировать",
                                      command=self.convert_file, width=20)
        self.convert_btn.grid(row=6, column=0, columnspan=3, pady=(0, 10))

        # Статусная строка
        self.status_var = tk.StringVar()
        self.status_var.set("Готов к работе")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var,
                                      font=('Arial', 9), relief=tk.SUNKEN,
                                      anchor=tk.W, padding=(5, 2))
        self.status_label.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E),
                               pady=(20, 0))

        # Настройка drag-and-drop для файлов
        self.setup_drag_drop()

    def setup_drag_drop(self):
        """Настройка перетаскивания файлов"""
        # Для Windows
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD
            # Если установлен tkinterdnd2, используем его
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.drop_file)
        except ImportError:
            pass  # Пропускаем, если tkinterdnd2 не установлен

    def drop_file(self, event):
        """Обработчик перетаскивания файла"""
        file_path = event.data
        # Удаляем фигурные скобки для Windows
        file_path = file_path.strip('{}')
        self.path_var.set(file_path)

    def browse_file(self):
        """Открыть диалог выбора файла"""
        file_path = filedialog.askopenfilename(
            title="Выберите CSV файл",
            filetypes=[("CSV файлы", "*.csv"), ("Все файлы", "*.*")]
        )
        if file_path:
            self.path_var.set(file_path)

    def convert_file(self):
        """Выполнить конвертацию файла"""
        # Получаем путь к файлу
        file_path = self.path_var.get().strip()

        if not file_path:
            messagebox.showerror("Ошибка", "Пожалуйста, укажите путь к CSV файлу")
            self.status_var.set("Ошибка: не указан путь к файлу")
            return

        # Получаем выбранный разделитель
        selected_sep_name = self.sep_var.get()
        separator = self.separator_map[selected_sep_name]

        # Обновляем статус
        self.status_var.set("Конвертация...")
        self.root.update()

        # Выполняем конвертацию
        result, info = converter(file_path, separator)

        if result == 0:
            self.status_var.set(f"Успешно! Файл сохранен: {info}")
            messagebox.showinfo("Успех",
                                f"Файл успешно конвертирован!\n\n"
                                f"Исходный файл: {file_path}\n"
                                f"Целевой разделитель: {selected_sep_name}\n"
                                f"Результат: {info}")
        elif result == 1:
            self.status_var.set("Ошибка: файл не найден")
            messagebox.showerror("Ошибка",
                                 f"Файл не найден!\n\n"
                                 f"Проверьте путь: {file_path}")
        else:
            self.status_var.set(f"Ошибка при конвертации: {info}")
            messagebox.showerror("Ошибка",
                                 f"Произошла ошибка при конвертации:\n\n{info}")


def main():
    root = tk.Tk()

    # Попытка использовать TkinterDnD для drag-and-drop
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        pass  # Используем обычный Tk

    app = CSVConverterApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()