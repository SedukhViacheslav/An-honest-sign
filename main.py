import csv
from collections import defaultdict
from tkinter import Tk, Label, Frame, messagebox, filedialog
from tkinter.ttk import Style, Button


def extract_gtin(gtin):
    """Извлекает штрих-код, обрезая ведущие нули."""
    return gtin.lstrip('0') or '0'


def process_file(input_file, output_file):
    """Обрабатывает CSV-файл и сохраняет дату, время, GTIN + статистику."""
    try:
        gtin_entries = []
        gtin_counts = defaultdict(int)
        total_gtins = 0

        with open(input_file, mode='r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                gtin = row.get('GTIN', '').strip()
                if not gtin:
                    continue

                operation_datetime = row.get(
                    'Дата и время выполнения операции, в результате которой было выявлено отклонение', '')
                if not operation_datetime:
                    continue

                date, time = operation_datetime.split(' ') if ' ' in operation_datetime else ('', '')
                cleaned_gtin = extract_gtin(gtin)

                gtin_entries.append((date, time, cleaned_gtin))
                gtin_counts[cleaned_gtin] += 1
                total_gtins += 1

        sorted_gtins = sorted(gtin_counts.items(), key=lambda x: -x[1])

        with open(output_file, mode='w', encoding='utf-8') as file:
            file.write("Дата       | Время   | Штрих-код\n")
            file.write("--------------------------------\n")
            for date, time, gtin in gtin_entries:
                file.write(f"{date} | {time} | {gtin}\n")

            file.write("\n\n=== Статистика GTIN ===\n")
            file.write("Штрих-код   | Количество\n")
            file.write("----------------------\n")
            for gtin, count in sorted_gtins:
                file.write(f"{gtin.ljust(12)} | {count}\n")

            file.write(f"\nВсего GTIN: {total_gtins}\n")
            file.write(f"Уникальных GTIN: {len(gtin_counts)}\n")

        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")
        return False


def center_window(window):
    """Центрирует окно на экране."""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')


def create_main_window():
    """Создаёт главное окно приложения."""
    root = Tk()
    root.title("Анализатор GTIN")
    root.configure(bg='#FFA500')  # Оранжевый фон

    # Стиль для кнопок
    style = Style()
    style.configure('TButton',
                    foreground='black',
                    background='#4CAF50',  # Зелёный цвет кнопок
                    font=('Arial', 10),
                    padding=5)

    # Основной фрейм
    frame = Frame(root, bg='#FFA500')
    frame.pack(pady=20, padx=20)

    # Информация о правообладателе
    Label(frame,
          text="Правообладатель: Седых Вячеслав\nEmail: Sedukh@ya.ru",
          bg='#FFA500',
          font=('Arial', 10)).pack(pady=10)

    # Кнопка выбора файла
    Button(frame,
           text="Выбрать файл CSV",
           command=lambda: select_and_process_file(root),
           style='TButton').pack(pady=10)

    center_window(root)
    root.mainloop()


def select_and_process_file(root):
    """Обрабатывает выбор файла и места сохранения."""
    input_file = filedialog.askopenfilename(
        title="Выберите файл CSV для анализа",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )

    if not input_file:
        return

    output_file = filedialog.asksaveasfilename(
        title="Выберите место для сохранения результата",
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        initialfile="result.txt"
    )

    if output_file:
        if process_file(input_file, output_file):
            messagebox.showinfo("Успех", f"Результат успешно сохранён в:\n{output_file}")
        root.destroy()


if __name__ == "__main__":
    create_main_window()