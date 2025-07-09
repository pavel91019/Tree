import pandas as pd
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox


def select_excel_file():
    root = tk.Tk()
    root.withdraw()  # Скрываем основное окно

    file_path = filedialog.askopenfilename(
        title="Выберите Excel файл с кодами",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )

    if not file_path:
        messagebox.showinfo("Информация", "Файл не выбран!")
        return None

    return file_path


def process_codes(codes):
    # Разбиваем коды на части по дефисам
    split_codes = [code.split('-') for code in codes]

    # Строим дерево
    tree = defaultdict(dict)

    for parts in split_codes:
        current_level = tree
        for i in range(len(parts)):
            # Для каждого уровня берем только соответствующую часть кода
            level_part = parts[i]
            if level_part not in current_level:
                current_level[level_part] = defaultdict(dict)
            current_level = current_level[level_part]

    return tree


def print_tree(node, level=0, prefix="", result=None):
    if result is None:
        result = []

    for part, child in sorted(node.items()):
        # Формируем отступы в зависимости от уровня
        indent = '  ' * level
        # Для уровня 1 выводим полностью первую часть (например "ФЕРм")
        if level == 0:
            display_part = part
        else:
            # Для остальных уровней выводим только текущую часть
            display_part = part.split('-')[-1] if '-' in part else part

        result.append(f"{indent}{display_part} (уровень {level + 1})")
        print_tree(child, level + 1, f"{prefix}-{part}" if prefix else part, result)

    return result


def main():
    file_path = select_excel_file()
    if not file_path:
        return

    try:
        # Загрузка данных из Excel (предполагаем, что данные в первом столбце)
        df = pd.read_excel(file_path, header=None)
        codes = df[0].astype(str).tolist()

        tree = process_codes(codes)
        result = print_tree(tree)

        # Создаем окно для отображения результатов
        result_window = tk.Tk()
        result_window.title("Результат группировки кодов")

        text = tk.Text(result_window, wrap=tk.WORD, width=60, height=30)
        scroll = tk.Scrollbar(result_window, command=text.yview)
        text.configure(yscrollcommand=scroll.set)

        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text.pack(fill=tk.BOTH, expand=True)

        text.insert(tk.END, "\n".join(result))
        text.config(state=tk.DISABLED)

        result_window.mainloop()

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка при обработке файла:\n{str(e)}")


if __name__ == "__main__":
    main()