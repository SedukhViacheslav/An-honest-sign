import csv
from collections import defaultdict
from tkinter import filedialog, Tk


def extract_gtin(gtin):
    """Извлекает штрих-код, обрезая ведущие нули."""
    return gtin.lstrip('0') or '0'  # Если после обрезки остаётся пусто, возвращаем '0'


def process_file(input_file, output_file):
    """Обрабатывает CSV-файл и сохраняет результат подсчёта GTIN."""
    gtin_counts = defaultdict(int)

    with open(input_file, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            gtin = row.get('GTIN', '').strip()
            if gtin:
                cleaned_gtin = extract_gtin(gtin)
                gtin_counts[cleaned_gtin] += 1

    # Сортируем по убыванию количества
    sorted_gtins = sorted(gtin_counts.items(), key=lambda x: -x[1])

    # Сохраняем результат в файл
    with open(output_file, mode='w', encoding='utf-8') as file:
        file.write("Штрих-код - Количество\n")
        file.write("----------------------\n")
        for gtin, count in sorted_gtins:
            file.write(f"{gtin} - {count}\n")

    print(f"Результат сохранён в {output_file}")


def main():
    """Запуск программы с выбором файла через диалог."""
    root = Tk()
    root.withdraw()  # Скрываем основное окно Tkinter

    input_file = filedialog.askopenfilename(
        title="Выберите файл CSV",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )

    if not input_file:
        print("Файл не выбран. Программа завершена.")
        return

    output_file = "result.txt"
    process_file(input_file, output_file)


if __name__ == "__main__":
    main()