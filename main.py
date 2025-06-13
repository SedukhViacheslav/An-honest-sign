import csv
from collections import defaultdict
from tkinter import Tk, Label, Frame, messagebox, filedialog, simpledialog
from tkinter.ttk import Style, Button
from datetime import datetime
import sys
import os
import re


def extract_gtin(gtin):
    """Извлекает штрих-код, обрезая ведущие нули."""
    return gtin.lstrip('0') or '0'


def resource_path(relative_path):
    """ Получает абсолютный путь к ресурсу, работает для dev и для PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def parse_microinvest_logs(log_files):
    """Парсит файлы логов Микроинвест (HTM) и извлекает информацию о марках."""
    mark_data = {}

    for log_file in log_files:
        try:
            with open(log_file, 'r', encoding='utf-8') as file:
                content = file.read()

                # Ищем все марки и даты в HTM-логе
                pattern = r'WHPL Request: \["([A-Za-z0-9+/=]+)"\].*?Date/Time: (\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2})'
                matches = re.findall(pattern, content, re.DOTALL)

                for mark_base64, check_time in matches:
                    try:
                        # Декодируем base64 (упрощенный пример)
                        import base64
                        mark_decoded = base64.b64decode(mark_base64).decode('utf-8')

                        # Добавляем в словарь (используем оригинальный код марки как ключ)
                        mark_data[mark_decoded] = check_time
                    except:
                        continue
        except Exception as e:
            print(f"Ошибка при чтении файла {log_file}: {str(e)}")

    return mark_data


def process_file(input_file, output_file, mark_data):
    """Обрабатывает CSV-файл и генерирует расширенный отчет с учетом данных из логов."""
    try:
        violations = defaultdict(list)
        general_info = {
            'Субъект': '',
            'Адрес': '',
            'ИНН участника': '',
            'Номер фискального накопителя': '',
            'Товарная группа': ''
        }

        with open(input_file, mode='r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if not general_info['Субъект']:
                    general_info = {
                        'Субъект': row.get('Субъект', ''),
                        'Адрес': row.get('Адрес места фиксации отклонения', ''),
                        'ИНН участника': row.get('ИНН участника', ''),
                        'Номер фискального накопителя': row.get('Фискальный номер накопителя из чека операции',
                                                                '') or row.get('Номер фискального накопителя', ''),
                        'Товарная группа': row.get('Товарная группа', '')
                    }

                violation_type = row.get('Вид отклонения', '')
                violation_desc = violation_type.split(':')[-1].strip() if ':' in violation_type else violation_type
                result = row.get('Результат проверки', '')
                gtin = extract_gtin(row.get('GTIN', '').strip())
                mark = row.get('Код', '').strip()
                receipt_number = row.get('Номер документа', '').strip()

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
                    # Проверяем наличие марки в логах Микроинвест
                    mark_status = "нет"
                    check_time = ""

                    # Ищем полное совпадение марки
                    if mark in mark_data:
                        mark_status = "да"
                        check_time = mark_data[mark]
                    else:
                        # Дополнительная проверка частичного совпадения
                        for log_mark, log_time in mark_data.items():
                            if mark in log_mark:
                                mark_status = "да (частичное)"
                                check_time = log_time
                                break

                    violations[(violation_desc, result)].append(
                        (date, time_str, gtin, mark, receipt_number, mark_status, check_time))

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

            markers = ['а)', 'б)', 'в)', 'г)', 'д)', 'е)', 'ж)', 'з)', 'и)', 'к)']
            for i, ((violation_desc, result), records) in enumerate(violations.items()):
                if i < len(markers):
                    marker = markers[i]
                else:
                    marker = f"{i + 1})"

                file.write(f"\n{marker} {violation_desc}\n")

                unique_gtins = set()
                dates = []
                for date, time_str, gtin, mark, receipt_number, mark_status, check_time in records:
                    unique_gtins.add(gtin)
                    dates.append(date)

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
                for date, time_str, gtin, mark, receipt_number, mark_status, check_time in sorted(records):
                    line = f"   * {date.strftime('%d.%m.%Y')} {time_str} - GTIN: {gtin.ljust(14)}"
                    line += f"\tМаркировка: {mark.ljust(25)}"
                    line += f"\tЧек №: {receipt_number.ljust(15)}"
                    line += f"\tВ логах: {mark_status}"
                    if "да" in mark_status:
                        line += f" (проверено: {check_time})"
                    file.write(line + "\n")

            # 3. Общая статистика
            file.write("\n\n" + "=" * 50 + "\n")
            file.write("ОБЩАЯ СТАТИСТИКА\n")
            file.write("=" * 50 + "\n")
            file.write(f"- Всего нарушений: {sum(len(v) for v in violations.values())}\n")
            file.write(
                f"- Уникальных GTIN: {len(set(gtin for records in violations.values() for _, _, gtin, _, _, _, _ in records))}\n")
            file.write(f"- Типов нарушений: {len(violations)}\n")
            found_in_logs = sum(1 for records in violations.values() for *_, status, _ in records if "да" in status)
            file.write(
                f"- Найдено в логах Микроинвест: {found_in_logs} из {sum(len(v) for v in violations.values())}\n")

        return True

    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка обработки файла: {str(e)}")
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
    root.title("Анализатор нарушений ЧЗ + проверка логов")
    root.configure(bg='#FFA500')

    try:
        root.iconbitmap(resource_path('icon.ico'))
    except:
        pass

    style = Style()
    style.configure('TButton',
                    foreground='black',
                    background='#4CAF50',
                    font=('Arial', 10),
                    padding=5)

    frame = Frame(root, bg='#FFA500')
    frame.pack(pady=20, padx=20)

    Label(frame,
          text="Анализ ошибок ЧЗ с проверкой логов Микроинвест",
          bg='#FFA500',
          font=('Arial', 10, 'bold')).pack(pady=10)

    Label(frame,
          text="Шаг 1: Выберите файлы логов (HTM)\nШаг 2: Выберите файл отчета ЧЗ (CSV)",
          bg='#FFA500',
          font=('Arial', 9)).pack(pady=5)

    Button(frame,
           text="Начать анализ",
           command=lambda: select_and_process_file(root),
           style='TButton').pack(pady=15)

    center_window(root)
    root.mainloop()


def select_and_process_file(root):
    """Обрабатывает выбор файлов и места сохранения."""
    # 1. Запрашиваем файлы логов Микроинвест
    log_files = filedialog.askopenfilenames(
        title="ШАГ 1: Выберите файлы логов Микроинвест (HTM)",
        filetypes=[("HTM logs", "*.htm"), ("All files", "*.*")],
        initialdir=os.getcwd()
    )

    if not log_files:
        if not messagebox.askyesno("Подтверждение",
                                   "Логи Микроинвест не выбраны. Продолжить без проверки логов?",
                                   icon='warning'):
            return

    # 2. Парсим логи
    mark_data = parse_microinvest_logs(log_files) if log_files else {}
    if log_files:
        messagebox.showinfo("Информация",
                            f"Загружено {len(mark_data)} марок из {len(log_files)} файлов логов")

    # 3. Запрашиваем файл CSV для анализа
    input_file = filedialog.askopenfilename(
        title="ШАГ 2: Выберите файл отчета ЧЗ (CSV)",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        initialdir=os.getcwd()
    )

    if not input_file:
        return

    # 4. Запрашиваем место сохранения отчета
    output_file = filedialog.asksaveasfilename(
        title="ШАГ 3: Сохранить отчет как...",
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        initialfile="Отчет_о_нарушениях.txt",
        initialdir=os.getcwd()
    )

    if output_file:
        if process_file(input_file, output_file, mark_data):
            messagebox.showinfo("Готово",
                                f"Отчет успешно сохранён:\n{output_file}\n\n"
                                f"Проверено марок в логах: {len(mark_data)}")
        root.destroy()


if __name__ == "__main__":
    create_main_window()