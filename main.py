import csv
from collections import defaultdict
from tkinter import Tk, Label, Frame, messagebox, filedialog
from tkinter.ttk import Style, Button
from datetime import datetime


def extract_gtin(gtin):
    """Извлекает штрих-код, обрезая ведущие нули."""
    return gtin.lstrip('0') or '0'


def process_file(input_file, output_file):
    """Обрабатывает CSV-файл и генерирует расширенный отчет."""
    try:
        violations = defaultdict(list)
        general_info = {
            'Субъект': '',
            'Адрес': '',
            'ИНН участника': '',
            'Номер фискального накопителя': '',  # Изменено название поля
            'Товарная группа': ''
        }

        with open(input_file, mode='r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                # Собираем общую информацию (возьмем из первой строки)
                if not general_info['Субъект']:
                    general_info = {
                        'Субъект': row.get('Субъект', ''),
                        'Адрес': row.get('Адрес места фиксации отклонения', ''),
                        'ИНН участника': row.get('ИНН участника', ''),
                        'Номер фискального накопителя': row.get('Фискальный номер накопителя из чека операции', ''),
                        # Берем из соответствующего поля
                        'Товарная группа': row.get('Товарная группа', '')
                    }

                violation_type = row.get('Вид отклонения', '')
                # Берем только текст после двоеточия в нарушении
                violation_desc = violation_type.split(':')[-1].strip() if ':' in violation_type else violation_type
                result = row.get('Результат проверки', '')
                gtin = extract_gtin(row.get('GTIN', '').strip())

                operation_datetime = row.get(
                    'Дата и время выполнения операции, в результате которой было выявлено отклонение', '')
                try:
                    dt = datetime.strptime(operation_datetime, '%Y-%m-%d %H:%M:%S')
                    date = dt.date()
                    time_str = dt.strftime('%H:%M')
                except:
                    date = None
                    time_str = ''

                if violation_desc and result and gtin and date:
                    violations[(violation_desc, result)].append((date, time_str, gtin))

        # Генерируем отчет
        with open(output_file, mode='w', encoding='utf-8') as file:
            # 1. Общие сведения
            file.write("=" * 50 + "\n")
            file.write("ОБЩИЕ СВЕДЕНИЯ\n")
            file.write("=" * 50 + "\n")
            for key, value in general_info.items():
                file.write(f"- {key}: {value}\n")

            # 2. Типы нарушений
            file.write("\n\n" + "=" * 50 + "\n")
            file.write("АНАЛИЗ НАРУШЕНИЙ\n")
            file.write("=" * 50 + "\n")

            # Используем буквенные маркеры для каждого типа нарушения
            markers = ['а)', 'б)', 'в)', 'г)', 'д)', 'е)', 'ж)', 'з)', 'и)', 'к)']
            for i, ((violation_desc, result), records) in enumerate(violations.items()):
                if i < len(markers):
                    marker = markers[i]
                else:
                    marker = f"{i + 1})"

                file.write(f"\n{marker} {violation_desc}\n")

                # Собираем уникальные GTIN и даты
                unique_gtins = set()
                dates = []
                for date, time_str, gtin in records:
                    unique_gtins.add(gtin)
                    dates.append(date)

                # Форматируем даты
                if dates:
                    min_date = min(dates)
                    max_date = max(dates)
                    date_range = f"{min_date.strftime('%d.%m.%Y')}-{max_date.strftime('%d.%m.%Y')}"
                else:
                    date_range = "нет данных"

                file.write(f" - Количество записей: {len(records)} случаев\n")
                file.write(f" - Период нарушений: {date_range}\n")
                file.write(f" - Уникальные штрих-коды (GTIN): {', '.join(sorted(unique_gtins))}\n")
                file.write(f" - Подробная информация:\n")
                for date, time_str, gtin in sorted(records):
                    file.write(f"   * {date.strftime('%d.%m.%Y')} {time_str} - {gtin}\n")

            # 3. Общая статистика
            file.write("\n\n" + "=" * 50 + "\n")
            file.write("ОБЩАЯ СТАТИСТИКА\n")
            file.write("=" * 50 + "\n")
            file.write(f"- Всего нарушений: {sum(len(v) for v in violations.values())}\n")
            file.write(
                f"- Уникальных GTIN: {len(set(gtin for records in violations.values() for _, _, gtin in records))}\n")
            file.write(f"- Типов нарушений: {len(violations)}\n")

        return True

    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка обработки файла: {str(e)}")
        return False


# Остальной код без изменений
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
    root.title("Анализатор нарушений")
    root.configure(bg='#FFA500')  # Оранжевый фон

    # Стиль для кнопок
    style = Style()
    style.configure('TButton',
                    foreground='black',
                    background='#4CAF50',
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
           text="Выбрать файл для анализа",
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
        title="Выберите место для сохранения отчета",
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        initialfile="Отчет_о_нарушениях.txt"
    )

    if output_file:
        if process_file(input_file, output_file):
            messagebox.showinfo("Успех", f"Отчет успешно сохранён в:\n{output_file}")
        root.destroy()


if __name__ == "__main__":
    create_main_window()