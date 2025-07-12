import pandas as pd
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class CodeProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализатор кодов ФЕР")
        self.root.geometry("800x600")

        self.max_level = tk.IntVar(value=4)  # Макс. уровень анализа
        self.min_collapse_level = tk.IntVar(value=3)  # Мин. уровень для свёртки (новое поле)
        self.file_path = None

        self.create_widgets()

    def create_widgets(self):
        # Фрейм для настроек
        settings_frame = ttk.LabelFrame(self.root, text="Настройки", padding=10)
        settings_frame.pack(fill=tk.X, padx=5, pady=5)

        # Поле для максимального уровня
        ttk.Label(settings_frame, text="Максимальный уровень:").grid(row=0, column=0, sticky=tk.W)
        self.level_spinbox = ttk.Spinbox(
            settings_frame,
            from_=1,
            to=10,
            textvariable=self.max_level,
            width=5
        )
        self.level_spinbox.grid(row=0, column=1, sticky=tk.W, padx=5)

        # Поле для минимального уровня свёртки (НОВОЕ)
        ttk.Label(settings_frame, text="Мин. уровень свёртки:").grid(row=1, column=0, sticky=tk.W)
        self.min_level_spinbox = ttk.Spinbox(
            settings_frame,
            from_=1,
            to=10,
            textvariable=self.min_collapse_level,
            width=5
        )
        self.min_level_spinbox.grid(row=1, column=1, sticky=tk.W, padx=5)

        # Кнопка выбора файла
        self.file_btn = ttk.Button(
            settings_frame,
            text="Выбрать файл Excel",
            command=self.select_file
        )
        self.file_btn.grid(row=2, column=0, columnspan=2, pady=5)

        # Кнопка обработки
        self.process_btn = ttk.Button(
            settings_frame,
            text="Обработать коды",
            command=self.process_codes,
            state=tk.DISABLED
        )
        self.process_btn.grid(row=3, column=0, columnspan=2, pady=5)

        # Кнопки управления деревом
        control_frame = ttk.Frame(settings_frame)
        control_frame.grid(row=4, column=0, columnspan=2, pady=5)

        ttk.Button(
            control_frame,
            text="Развернуть все",
            command=self.expand_all
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(
            control_frame,
            text="Свернуть все",
            command=self.collapse_all
        ).pack(side=tk.LEFT, padx=2)

        # Фрейм для результатов
        result_frame = ttk.LabelFrame(self.root, text="Результаты", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Treeview с прокруткой
        self.tree = ttk.Treeview(result_frame, show="tree")
        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(
            title="Выберите Excel файл с кодами",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if self.file_path:
            self.process_btn.config(state=tk.NORMAL)
        else:
            self.process_btn.config(state=tk.DISABLED)

    def process_codes(self):
        if not self.file_path:
            messagebox.showwarning("Предупреждение", "Сначала выберите файл!")
            return

        try:
            # Загрузка данных из Excel
            df = pd.read_excel(self.file_path, header=None)

            # Собираем все коды
            all_codes = []
            for cell in df[0]:
                if pd.notna(cell):
                    codes_in_cell = str(cell).split('\n')
                    for code in codes_in_cell:
                        code = code.strip()
                        if code:
                            all_codes.append(code)

            # Очистка дерева
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Построение дерева
            self.build_tree(all_codes)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")

    def build_tree(self, codes):
        tree_dict = defaultdict(dict)
        max_level = self.max_level.get()

        for code in codes:
            parts = code.split('-')
            current_level = tree_dict

            for i in range(min(len(parts), max_level)):
                part = parts[i].strip()
                if part not in current_level:
                    current_level[part] = defaultdict(dict)
                current_level = current_level[part]

        # Добавление в Treeview
        self.add_tree_items("", tree_dict, level=1)

    def add_tree_items(self, parent, node, level):
        for part, child in sorted(node.items()):
            item = self.tree.insert(parent, "end", text=part, open=level < self.min_collapse_level.get())
            if child:
                self.add_tree_items(item, child, level + 1)

    def expand_all(self):
        for item in self.tree.get_children():
            self.expand_item_and_children(item, level=1)

    def expand_item_and_children(self, item, level):
        if level >= self.min_collapse_level.get():
            self.tree.item(item, open=True)
        for child in self.tree.get_children(item):
            self.expand_item_and_children(child, level + 1)

    def collapse_all(self):
        for item in self.tree.get_children():
            self.collapse_item_and_children(item, level=1)

    def collapse_item_and_children(self, item, level):
        if level >= self.min_collapse_level.get():
            self.tree.item(item, open=False)
        for child in self.tree.get_children(item):
            self.collapse_item_and_children(child, level + 1)


if __name__ == "__main__":
    root = tk.Tk()
    app = CodeProcessorApp(root)
    root.mainloop()