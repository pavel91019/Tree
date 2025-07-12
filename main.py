import pandas as pd
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json


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

        # Кнопки сохранения/загрузки
        btn_frame = ttk.Frame(settings_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        ttk.Button(
            btn_frame,
            text="Сохранить дерево",
            command=self.save_tree
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        ttk.Button(
            btn_frame,
            text="Загрузить дерево",
            command=self.load_tree
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

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

    def save_tree(self):
        """Сохраняет дерево с раскраской в файл"""
        if not self.tree.get_children():
            messagebox.showwarning("Внимание", "Дерево пустое, нечего сохранять")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Сохранить дерево"
        )

        if not file_path:
            return

        try:
            # Собираем данные для сохранения
            tree_data = {
                "structure": self.get_tree_structure(),
                "colored_items": list(self.colored_items),
                "current_level": self.current_level.get()
            }

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(tree_data, f, ensure_ascii=False, indent=2)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить дерево:\n{str(e)}")

    def get_tree_structure(self):
        """Рекурсивно собирает структуру дерева"""
        def get_children(parent=""):
            children = []
            for item in self.tree.get_children(parent):
                children.append({
                    "text": self.tree.item(item, "text"),
                    "open": self.tree.item(item, "open"),
                    "children": get_children(item)
                })
            return children

        return get_children()

    def load_tree(self):
        """Загружает дерево из файла"""
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Загрузить дерево"
        )

        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                tree_data = json.load(f)

            # Очищаем текущее дерево
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.colored_items.clear()

            # Восстанавливаем структуру
            self.rebuild_tree(tree_data["structure"])

            # Восстанавливаем раскраску
            self.colored_items = set(tree_data["colored_items"])
            for item in self.colored_items:
                if self.tree.exists(item):  # Проверяем, что элемент существует
                    self.tree.item(item, tags=('green',))

            # Восстанавливаем уровень отображения
            self.current_level.set(tree_data.get("current_level", 3))
            self.update_tree_display()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить дерево:\n{str(e)}")

    def rebuild_tree(self, structure, parent=""):
        """Рекурсивно восстанавливает структуру дерева"""
        for item_data in structure:
            item = self.tree.insert(
                parent,
                "end",
                text=item_data["text"],
                open=item_data["open"]
            )
            if item_data["children"]:
                self.rebuild_tree(item_data["children"], item)


if __name__ == "__main__":
    root = tk.Tk()
    app = CodeProcessorApp(root)
    root.mainloop()