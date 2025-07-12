import pandas as pd
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class CodeProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализатор кодов ФЕР")
        self.root.geometry("800x600")

        self.current_level = tk.IntVar(value=3)
        self.file_path = None
        self.colored_items = set()  # Множество для хранения окрашенных элементов

        self.create_widgets()
        self.setup_styles()

    def setup_styles(self):
        def setup_styles(self):
            self.style = ttk.Style()
            # Настраиваем стиль для элементов с тегом 'green'
            self.style.configure("Treeview", background="white")
            self.style.map("Treeview",
                           background=[('selected', '#347083')])  # Цвет выделения
            self.style.configure("Treeview.green", background="#e6ffe6")
            self.style.map("Treeview.green",
                           background=[('selected', '#88cce6')])  # Цвет выделения для зеленых элементов

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

        # Кнопка раскрашивания/снятия цвета
        self.color_btn = ttk.Button(
            settings_frame,
            text="Переключить цвет выделенного",
            command=self.toggle_color_selected
        )
        self.color_btn.pack(fill=tk.X, pady=5)

        # Фрейм для результатов
        result_frame = ttk.LabelFrame(self.root, text="Результаты", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Treeview с прокруткой
        self.tree = ttk.Treeview(result_frame, show="tree", selectmode="extended")
        self.tree.tag_configure('green', background='#e6ffe6')  # Конфигурация тега
        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)
    def select_file(self):
        self.file_path = filedialog.askopenfilename(
            title="Выберите Excel файл с кодами",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

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

            # Очистка дерева и сброс цветов
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.colored_items.clear()

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
        if 1 <= new_level <= 10:
            self.current_level.set(new_level)
            self.update_tree_display()

    def update_tree_display(self):
        target_level = self.current_level.get()

        def set_item_state(item, level):
            is_open = level < target_level
            self.tree.item(item, open=is_open)

            # Сохраняем цвет при обновлении отображения
            if item in self.colored_items:
                self.tree.item(item, tags=("green",))
            else:
                self.tree.item(item, tags=())

            for child in self.tree.get_children(item):
                set_item_state(child, level + 1)

        for item in self.tree.get_children():
            set_item_state(item, 1)

    def toggle_color_selected(self):
        """Переключает цвет выделенных элементов"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Внимание", "Не выбрано ни одного элемента")
            return

        for item in selected_items:
            if item in self.colored_items:
                # Если элемент уже окрашен - убираем цвет
                self.colored_items.remove(item)
                self.tree.item(item, tags=())  # Удаляем тег
            else:
                # Если элемент не окрашен - добавляем цвет
                self.colored_items.add(item)
                self.tree.item(item, tags=('green',))  # Применяем тег

if __name__ == "__main__":
    root = tk.Tk()
    app = CodeProcessorApp(root)
    root.mainloop()