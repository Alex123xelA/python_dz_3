import os
import pandas as pd


def select_dataframe() -> pd.DataFrame:
    """
    Позволяет пользователю выбрать один из .pkl-файлов и загружает DataFrame.
    """
    files = [f for f in os.listdir("./data") if f.endswith(".pkl")]
    if not files:
        raise FileNotFoundError("Нет доступных .pkl файлов в папке ./data")

    print("Доступные справочники:")
    for i, f in enumerate(files):
        print(f"{i + 1}. {f}")

    index = int(input("Выберите справочник по номеру: ")) - 1
    if not (0 <= index < len(files)):
        raise IndexError("Некорректный номер справочника")

    df = pd.read_pickle(os.path.join("./data", files[index]))
    df.columns = df.columns.str.strip()

    print("\nКолонки:")
    for i, col in enumerate(df.columns):
        print(f"{i + 1}. {col}")
    print("\nПервые строки таблицы:")
    print(df.head())
    return df


def load_excel_to_pickle(excel_path: str = 'DZ_2.xlsx', output_dir: str = './data/') -> None:
    """
    Загружает все листы Excel-файла (кроме первого) и сохраняет каждый как .pkl-файл.
    """
    os.makedirs(output_dir, exist_ok=True)

    xl = pd.ExcelFile(excel_path)
    sheet_names = xl.sheet_names[1:]
    if not sheet_names:
        raise ValueError("Excel-файл не содержит листов кроме первого.")

    for sheet in sheet_names:
        df = xl.parse(sheet)
        df.columns = df.columns.str.strip()
        file_path = os.path.join(output_dir, f'{sheet}.pkl')
        df.to_pickle(file_path)
        print(f"[✓] Сохранено: {sheet}.pkl")


if __name__ == "__main__":
    load_excel_to_pickle()
