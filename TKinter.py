from tkinter import Tk, Label, Button, filedialog, messagebox, simpledialog, scrolledtext
from tkinter import ttk


class TKinter():
    def __init__(self):
        self.root = Tk()

        self.root.title('Делятор-V2')
        self.root.geometry('400x300')

        self.root.deiconify()  # Показать окно

        self.log_frame = ttk.Frame(self.root)
        self.log_frame.pack(pady=10)

        self.log_text = scrolledtext.ScrolledText(self.log_frame, wrap='word', width=50, height=10, state='disabled')
        self.log_text.pack()

        self.status_label = Label(self.root, text="Нажмите кнопку 'Выбрать файл' для деления.")
        self.status_label.pack(pady=10)

        # Кнопка для выбора файла
        self.process_button = Button(self.root, text="Выбрать файл", command=self.on_process_button)
        self.process_button.pack(pady=5)

        # Прогресс бар
        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=300, mode='determinate')
        self.progress_bar.pack(pady=10)

    def set_process_callback(self, callback):
        self.process_callback = callback

    def on_process_button(self):
        if hasattr(self, 'process_callback'):
            self.progress_bar['value'] = 0
            self.progress_bar['maximum'] = 100
            # Выполняем процесс сразу в основном потоке
            self.run_process()

    def run_process(self):
        try:
            # Запускаем основной процесс в основном потоке
            self.process_callback()
        except Exception as e:
            # Обработка ошибок, используя метод after для безопасности
            self.root.after(0, lambda err=e: messagebox.showerror("Ошибка", err))
        finally:
            # Завершаем прогресс-бар, также используя after для безопасности
            self.root.after(0, lambda: self.progress_bar.config(value=100))

    def update_progress(self, value):
        # Обновление прогресс-бара через after, чтобы избежать ошибок при работе с интерфейсом
        self.root.after(0, lambda: self.progress_bar.config(value=value))
        self.root.after(0, lambda: self.root.update_idletasks())

    def choose_file(self):
        file_path = filedialog.askopenfilename(title='Выберите файл Excel', filetypes=[("Excel files", "*.xlsx")])
        return file_path

    def choose_columns(self):
        def columns_letter_to_number(letter):
            try:
                return ord(letter.upper()) - 65
            except Exception:
                return None

        # Используем after, чтобы диалог также был безопасен
        genus_column_letter = simpledialog.askstring(title="Выбор столбца для папки",
                                                     prompt="Введите букву столбца для папки:")
        species_column_letter = simpledialog.askstring(title="Выбор столбца для файла",
                                                       prompt="Введите букву столбца для файла:")

        genus_column_num = columns_letter_to_number(genus_column_letter)
        species_column_num = columns_letter_to_number(species_column_letter)

        if genus_column_num is None or species_column_num is None:
            messagebox.showerror("Ошибка", "Введены некорректные буквы столбцов. Попробуйте снова.")
            return None, None

        return genus_column_num, species_column_num

    def log(self, message):
        # Для изменения интерфейса используем after
        self.root.after(0, lambda: self._log(message))

    def _log(self, message):
        self.log_text.config(state='normal')  # Разрешить редактирование
        self.log_text.insert('end', message + '\n')
        self.log_text.see('end')  # Прокрутка вниз
        self.log_text.config(state='disabled')  # Запретить редактирование

    def mainloop(self):
        self.root.mainloop()
