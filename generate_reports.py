import pandas as pd
import matplotlib.pyplot as plt

def choose_columns_by_index(columns: list, count: int = None) -> list:
    """
    Позволяет пользователю выбрать колонки по индексам.

    Parameters:
        columns (list): Список колонок.
        count (int): Ожидаемое количество колонок (или None для произвольного количества).

    Returns:
        list: Выбранные колонки.
    """
    for i, col in enumerate(columns):
        print(f"{i + 1}. {col}")
    raw = input("Введите номера колонок через запятую: ").split(",")

    selected = []
    for item in raw:
        item = item.strip()
        if item.isdigit():
            idx = int(item) - 1
            if 0 <= idx < len(columns):
                selected.append(columns[idx])
    if count and len(selected) != count:
        print(f"Ожидалось выбрать {count} колонок.")
        return []
    return selected


def generate_text_report(df: pd.DataFrame):
    """
    Генерирует текстовый отчет с фильтрацией по нескольким колонкам и выбором столбцов.
    """
    print("\nВыберите колонки для фильтрации:")
    filter_cols = choose_columns_by_index(df.columns.tolist())
    if not filter_cols:
        return

    criteria = {}
    for col in filter_cols:
        val = input(f"Введите значение для фильтрации по '{col}': ").strip()
        criteria[col] = val

    print("\nВыберите колонки для отображения в отчете:")
    selected_columns = choose_columns_by_index(df.columns.tolist())
    if not selected_columns:
        return

    mask = pd.Series(True, index=df.index)
    for col, val in criteria.items():
        mask &= df[col].astype(str) == val

    result = df.loc[mask, selected_columns]
    print("\nРезультат отчета:")
    if result.empty:
        print("Нет данных, соответствующих заданным фильтрам.")
    else:
        print(result)


def generate_scatter_plot(df: pd.DataFrame):
    """
    Строит scatter-график по любым двум колонкам (не обязательно числовым).
    """
    print("\nВыберите колонку X:")
    x_col = choose_columns_by_index(df.columns.tolist(), 1)
    if not x_col:
        return
    x = x_col[0]

    print("\nВыберите колонку Y:")
    y_col = choose_columns_by_index(df.columns.tolist(), 1)
    if not y_col:
        return
    y = y_col[0]

    try:
        df.plot.scatter(x=x, y=y)
        plt.title(f"{y} от {x}")
        plt.grid(True)
        plt.tight_layout()
        plt.show()
    except Exception as e:
        print(f"Ошибка при построении scatter-графика: {e}")


def generate_pie_chart(df: pd.DataFrame):
    """
    Строит круговую диаграмму по значениям в выбранной колонке.
    """
    print("\nВыберите колонку для круговой диаграммы:")
    selected = choose_columns_by_index(df.columns.tolist(), 1)
    if not selected:
        return
    col = selected[0]

    value_counts = df[col].value_counts()
    if value_counts.empty:
        print("Нет данных для диаграммы.")
        return

    plt.figure(figsize=(8, 8))
    value_counts.plot.pie(autopct='%1.1f%%', startangle=360, shadow=True)
    plt.title(f'Распределение по колонке: {col}')
    plt.ylabel('')
    plt.tight_layout()
    plt.show()


def generate_bar_chart(df: pd.DataFrame):
    """
    Строит столбчатую диаграмму по количеству значений в выбранной колонке.
    """
    print("\nВыберите колонку для столбчатой диаграммы:")
    selected = choose_columns_by_index(df.columns.tolist(), 1)
    if not selected:
        return
    col = selected[0]

    value_counts = df[col].value_counts()
    if value_counts.empty:
        print("Нет данных для диаграммы.")
        return

    value_counts.plot.bar()
    plt.title(f'Распределение по колонке: {col}')
    plt.ylabel("Количество")
    plt.tight_layout()
    plt.show()

def generate_pivot_report(df: pd.DataFrame) -> None:
    """
    Создает сводную таблицу с помощью pd.pivot_table.
    """
    print("\nВыберите параметры для сводной таблицы:")

    columns = df.columns.tolist()
    for i, col in enumerate(columns):
        print(f"{i + 1}. {col}")

    try:
        idx = int(input("Выберите колонку для строк (index): ")) - 1
        col = int(input("Выберите колонку для столбцов (columns): ")) - 1
        use_values = input("Хотите указать колонку для значений? (y/n): ").strip().lower() == "y"

        values = None
        if use_values:
            val = int(input("Выберите колонку для значений (values): ")) - 1
            values = columns[val]

        agg = input("Введите функцию агрегации (sum, count, mean, size и т.д.): ").strip()

        pivot = pd.pivot_table(
            df,
            index=columns[idx],
            columns=columns[col],
            values=values,
            aggfunc=agg,
            fill_value=0,
        )
        print("\nСводная таблица:")
        print(pivot)
    except Exception as e:
        print("Ошибка при создании сводной таблицы:", e)
