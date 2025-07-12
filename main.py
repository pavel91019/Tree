import pandas as pd
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class CodeProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализатор кодов ФЕР")
        self.root.geometry("800x600")

        self.current_level = tk.IntVar(value=3)  # Текущий уровень отображения
        self.file_path = None

        self.create_widgets()

    def create_widgets(self):
        # Фрейм для настроек
        settings_frame = ttk.LabelFrame(self.root, text="Настройки", padding=10)
        settings_frame.pack(fill=tk.X, padx=5, pady=5)

        # Поле для текущего уровня
        level_frame = ttk.Frame(settings_frame)
        level_frame.pack(fill=tk.X, pady=5)

        ttk.Label(level_frame, text="Текущий уровень:").pack(side=tk.LEFT)

        self.level_label = ttk.Label(level_frame, textvariable=self.current_level, width=3)
        self.level_label.pack(side=tk.LEFT, padx=5)

        # Кнопки управления уровнем
        ttk.Button(
            level_frame,
            text="↑ Увеличить",
            command=lambda: self.change_level(1)
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(
            level_frame,
            text="↓ Уменьшить",
            command=lambda: self.change_level(-1)
        ).pack(side=tk.LEFT, padx=2)

        # Кнопка выбора файла
        self.file_btn = ttk.Button(
            settings_frame,
            text="Выбрать файл Excel",
            command=self.select_file
        )
        self.file_btn.pack(fill=tk.X, pady=5)

        # Кнопка обработки
        self.process_btn = ttk.Button(
            settings_frame,
            text="Обработать коды",
            command=self.process_codes,
            state=tk.DISABLED
        )
        self.process_btn.pack(fill=tk.X, pady=5)

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
            self.update_tree_display()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")

    def build_tree(self, codes):
        tree_dict = defaultdict(dict)

        for code in codes:
            parts = code.split('-')
            current_level = tree_dict

            for part in parts:
                part = part.strip()
                if part not in current_level:
                    current_level[part] = defaultdict(dict)
                current_level = current_level[part]

        # Добавление в Treeview
        self.add_tree_items("", tree_dict, level=1)

    def add_tree_items(self, parent, node, level):
        for part, child in sorted(node.items()):
            item = self.tree.insert(parent, "end", text=part, open=False)
            if child:
                self.add_tree_items(item, child, level + 1)

    def change_level(self, delta):
        new_level = self.current_level.get() + delta
        if 1 <= new_level <= 10:  # Ограничиваем диапазон 1-10
            self.current_level.set(new_level)
            self.update_tree_display()

    def update_tree_display(self):
        target_level = self.current_level.get()

        def set_item_state(item, level):
            # Устанавливаем состояние узла (open/closed) в зависимости от уровня
            is_open = level < target_level
            self.tree.item(item, open=is_open)

            # Рекурсивно обрабатываем дочерние элементы
            for child in self.tree.get_children(item):
                set_item_state(child, level + 1)

        # Применяем ко всем корневым элементам
        for item in self.tree.get_children():
            set_item_state(item, 1)


if __name__ == "__main__":
    root = tk.Tk()
    app = CodeProcessorApp(root)
    root.mainloop()