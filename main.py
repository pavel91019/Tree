import pandas as pd
from collections import defaultdict


def load_and_process_excel(file_path):
    # Загрузка данных из Excel (предполагаем, что данные в первом столбце)
    df = pd.read_excel(file_path, header=None)
    codes = df[0].tolist()

    # Разбиваем коды на части по дефисам
    split_codes = [code.split('-') for code in codes]

    # Строим дерево
    tree = defaultdict(dict)

    for parts in split_codes:
        current_level = tree
        for i, part in enumerate(parts):
            prefix = '-'.join(parts[:i]) + '-' if i > 0 else ''
            full_part = prefix + part
            if full_part not in current_level:
                current_level[full_part] = defaultdict(dict)
            current_level = current_level[full_part]

    # Функция для рекурсивного вывода дерева
    def print_tree(node, level=0):
        for key, child in node.items():
            print('-' * level + key + f' (уровень {level + 1})')
            print_tree(child, level + 1)

    print_tree(tree)


# Пример использования
file_path = input("Введите путь к Excel файлу: ")
load_and_process_excel(file_path)