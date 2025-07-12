import pandas as pd
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class CodeProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализатор кодов ФЕР с чекбоксами")
        self.root.geometry("1000x800")

        # Состояния чекбоксов {item_id: checked}
        self.checked_items = {}

        self.current_level = tk.IntVar(value=3)
        self.file_path = None

        self.create_widgets()
        self.setup_styles()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.configure("Treeview", rowheight=25)
        self.style.map("Treeview.Heading", background=[('active', '#e6e6e6')])

    def create_widgets(self):
        # Main frames
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки", padding=10)
        settings_frame.pack(fill=tk.X, padx=5, pady=5)

        # Level control
        level_frame = ttk.Frame(settings_frame)
        level_frame.pack(fill=tk.X, pady=5)

        ttk.Label(level_frame, text="Уровень детализации:").pack(side=tk.LEFT)
        ttk.Label(level_frame, textvariable=self.current_level, width=2).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            level_frame,
            text="↑",
            command=lambda: self.change_level(1),
            width=3
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(
            level_frame,
            text="↓",
            command=lambda: self.change_level(-1),
            width=3
        ).pack(side=tk.LEFT, padx=2)

        # File buttons
        btn_frame = ttk.Frame(settings_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        ttk.Button(
            btn_frame,
            text="Выбрать файл",
            command=self.select_file
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(
            btn_frame,
            text="Сохранить отметки",
            command=self.save_checks
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(
            btn_frame,
            text="Загрузить отметки",
            command=self.load_checks
        ).pack(side=tk.LEFT, padx=2)

        # Treeview frame
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Treeview with checkboxes
        self.tree = ttk.Treeview(
            tree_frame,
            columns=("check", "item"),
            show="tree headings",
            selectmode="extended"
        )

        # Configure columns
        self.tree.column("#0", width=0, stretch=tk.NO)  # Hide first empty column
        self.tree.column("check", width=40, anchor="center")
        self.tree.column("item", width=900, anchor="w")

        # Headings
        self.tree.heading("check", text="✓", command=self.toggle_all)
        self.tree.heading("item", text="Код")

        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Bind events
        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<space>", self.toggle_selected)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if self.file_path:
            self.process_file()

    def process_file(self):
        try:
            df = pd.read_excel(self.file_path, header=None)
            all_codes = []

            for cell in df[0]:
                if pd.notna(cell):
                    codes_in_cell = str(cell).split('\n')
                    for code in codes_in_cell:
                        code = code.strip()
                        if code:
                            all_codes.append(code)

            self.build_tree(all_codes)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка обработки файла:\n{str(e)}")

    def build_tree(self, codes):
        # Clear existing tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.checked_items.clear()

        # Build tree structure
        tree_dict = defaultdict(dict)
        for code in codes:
            parts = code.split('-')
            current = tree_dict
            for part in parts:
                part = part.strip()
                if part not in current:
                    current[part] = defaultdict(dict)
                current = current[part]

        # Add items to treeview
        self.add_tree_items("", tree_dict, level=1)
        self.update_tree_display()

    def add_tree_items(self, parent, node, level):
        for part, child in sorted(node.items()):
            item_id = self.tree.insert(
                parent,
                "end",
                values=("☐", part),
                tags=(f"level{level}",)
            )

            if child:
                self.add_tree_items(item_id, child, level + 1)

    def update_tree_display(self):
        target_level = self.current_level.get()

        def set_visibility(item, level):
            visible = level <= target_level
            self.tree.item(item, open=level < target_level)

            for child in self.tree.get_children(item):
                set_visibility(child, level + 1)

        for item in self.tree.get_children():
            set_visibility(item, 1)

    def change_level(self, delta):
        new_level = self.current_level.get() + delta
        if 1 <= new_level <= 10:
            self.current_level.set(new_level)
            self.update_tree_display()

    def on_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        column = self.tree.identify_column(event.x)

        if region == "cell" and column == "#1":  # Checkbox column
            item = self.tree.identify_row(event.y)
            self.toggle_item(item)

    def toggle_item(self, item):
        current = self.tree.item(item, "values")[0]
        new_state = "☑" if current == "☐" else "☐"
        self.tree.item(item, values=(new_state, item.split("-")[-1]))

        # Update checked items dictionary
        self.checked_items[item] = (new_state == "☑")

        # Update parent states if needed
        self.update_parent_states(item)

    def toggle_selected(self, event):
        for item in self.tree.selection():
            self.toggle_item(item)

    def toggle_all(self):
        all_checked = all(
            self.tree.item(item, "values")[0] == "☑"
            for item in self.tree.get_children()
        )

        new_state = "☐" if all_checked else "☑"

        def set_state(item):
            self.tree.item(item, values=(new_state, self.tree.item(item, "values")[1]))
            self.checked_items[item] = (new_state == "☑")
            for child in self.tree.get_children(item):
                set_state(child)

        for item in self.tree.get_children():
            set_state(item)

    def update_parent_states(self, child_item):
        parent = self.tree.parent(child_item)
        if not parent:
            return

        children = self.tree.get_children(parent)
        if not children:
            return

        # Check if all siblings are checked
        all_checked = all(
            self.tree.item(item, "values")[0] == "☑"
            for item in children
        )

        # Check if any sibling is checked
        any_checked = any(
            self.tree.item(item, "values")[0] == "☑"
            for item in children
        )

        current_parent_value = self.tree.item(parent, "values")[0]

        if all_checked and current_parent_value != "☑":
            self.tree.item(parent, values=("☑", self.tree.item(parent, "values")[1]))
            self.checked_items[parent] = True
            self.update_parent_states(parent)
        elif not all_checked and any_checked and current_parent_value != "◐":
            self.tree.item(parent, values=("◐", self.tree.item(parent, "values")[1]))
            self.checked_items[parent] = None
            self.update_parent_states(parent)
        elif not any_checked and current_parent_value != "☐":
            self.tree.item(parent, values=("☐", self.tree.item(parent, "values")[1]))
            self.checked_items[parent] = False
            self.update_parent_states(parent)

    def save_checks(self):
        if not self.checked_items:
            messagebox.showinfo("Информация", "Нет отметок для сохранения")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if file_path:
            import json
            with open(file_path, "w") as f:
                json.dump(self.checked_items, f)
            messagebox.showinfo("Успех", "Отметки сохранены")

    def load_checks(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if file_path:
            import json
            try:
                with open(file_path, "r") as f:
                    checks = json.load(f)

                for item in self.tree.get_children(""):
                    self.apply_checks(item, checks)

                messagebox.showinfo("Успех", "Отметки загружены")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка загрузки:\n{str(e)}")

    def apply_checks(self, parent, checks):
        if parent in checks:
            state = "☑" if checks[parent] else "☐"
            self.tree.item(parent, values=(state, self.tree.item(parent, "values")[1]))
            self.checked_items[parent] = checks[parent]

        for child in self.tree.get_children(parent):
            self.apply_checks(child, checks)


if __name__ == "__main__":
    root = tk.Tk()
    app = CodeProcessorApp(root)
    root.mainloop()