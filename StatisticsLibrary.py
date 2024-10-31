import asyncio
from google.oauth2.service_account import Credentials
import gspread_asyncio

# Настройка OAuth scopes
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Функция для создания асинхронного клиента
def get_creds():
    return Credentials.from_service_account_file(
        'C:/Users/pasch/Пашино/Колледж/4 КУРС/Практика/mylibrarypython-3e7c1bcefbd1.json', scopes=SCOPES
    )

# Создание асинхронного клиента gspread
agcm = gspread_asyncio.AsyncioGspreadClientManager(get_creds)

# Метод для добавления нового студента
async def add_student(student_name, group, worksheet, absences=0):
    try:
        new_row = [student_name, group, absences if absences else '', '']
        await worksheet.append_row(new_row)
        print(f"Студент {student_name} добавлен в таблицу с группой {group}. Пропуски: {absences if absences else 'нет'}")
    except Exception as e:
        print(f"Ошибка при добавлении студента: {e}")

# Метод для добавления количества пропусков студенту
async def add_absences(student_name, absences, worksheet):
    try:
        student_cell = await worksheet.find(student_name)
        if not student_cell:
            print(f"Студент '{student_name}' не найден.")
            return

        absence_column = student_cell.col + 2  # 3-й столбец для "с 21 по 27 октября"
        total_absences_column = student_cell.col + 3  # 4-й столбец для "За полугодие"

        current_absences_cell = await worksheet.cell(student_cell.row, absence_column)
        current_absences = current_absences_cell.value

        if current_absences:
            current_absences = int(current_absences)
        else:
            current_absences = 0

        new_total_absences = current_absences + absences
        await worksheet.update_cell(student_cell.row, absence_column, new_total_absences)
        print(f"Добавлено {absences} пропусков для студента {student_name}. Всего пропусков за неделю: {new_total_absences}.")

        total_absences_cell = await worksheet.cell(student_cell.row, total_absences_column)
        total_absences = total_absences_cell.value

        if total_absences:
            total_absences = int(total_absences) + absences
        else:
            total_absences = absences

        await worksheet.update_cell(student_cell.row, total_absences_column, total_absences)
        print(f"Обновлено общее количество пропусков за полугодие для студента {student_name}: {total_absences}.")

    except Exception as e:
        print(f"Ошибка при добавлении пропусков: {e}")

# Метод для обновления заголовка даты
async def update_date_column_header(new_date_range, worksheet):
    try:
        absence_column = 3

        cells = await worksheet.get_all_values()
        for i in range(1, len(cells)):
            await worksheet.update_cell(i + 1, absence_column, '')

        await worksheet.update_cell(1, absence_column, new_date_range)
        print(f"Заголовок даты обновлен на: {new_date_range} и все пропуски очищены.")
    except Exception as e:
        print(f"Ошибка при обновлении заголовка: {e}")

# Метод для вывода полной статистики
async def print_full_statistics(worksheet):
    try:
        records = await worksheet.get_all_values()

        col_widths = [max(len(str(cell)) for cell in col) for col in zip(*records)]

        header = records[0]
        header_format = " | ".join(f"{header[i]:<{col_widths[i]}}" for i in range(len(header)))
        print(header_format)
        print("-" * len(header_format))

        for row in records[1:]:
            row_format = " | ".join(f"{row[i]:<{col_widths[i]}}" for i in range(len(row)))
            print(row_format)

    except Exception as e:
        print(f"Ошибка при выводе статистики: {e}")

# Основная программа для тестирования методов
async def main():
    client = await agcm.authorize()
    sheet = await client.open_by_key('1yB02wlZ6p12oM7hOvPAknFyC-aSUSaAoBJhB0Ru3VhE')
    worksheet = await sheet.get_worksheet(0)

    print("Выберите действие:")
    print("1. Поменять дату")
    print("2. Добавить пропуски")
    print("3. Вывести полную статистику")
    print("4. Добавить нового студента")

    choice = input("Введите номер действия (1, 2, 3 или 4): ").strip()

    if choice == '1':
        change_date = input("Поменять дату? (Да/Нет): ").strip().lower()
        if change_date == 'да':
            new_date_range = input("Введите новый диапазон дат (например, 'с 1 по 7 октября'): ").strip()
            await update_date_column_header(new_date_range, worksheet)

    elif choice == '2':
        add_absence = input("Добавить пропуски студенту? (Да/Нет): ").strip().lower()
        if add_absence == 'да':
            student_name = input("Введите ФИО студента для отметки пропусков: ").strip()
            absences = input(f"Введите количество пропусков для студента {student_name}: ").strip()

            if absences.isdigit():
                await add_absences(student_name, int(absences), worksheet)
            else:
                print("Ошибка: количество пропусков должно быть числом.")

    elif choice == '3':
        await print_full_statistics(worksheet)

    elif choice == '4':
        student_name = input("Введите ФИО нового студента: ").strip()
        group = input("Введите номер группы студента: ").strip()

        add_absence = input("Ввести пропуски для студента? (Да/Нет): ").strip().lower()
        if add_absence == 'да':
            absences = input(f"Введите количество пропусков для студента {student_name}: ").strip()
            if absences.isdigit():
                await add_student(student_name, group, worksheet, int(absences))
            else:
                print("Ошибка: количество пропусков должно быть числом.")
        else:
            await add_student(student_name, group, worksheet)

    else:
        print("Неверный выбор. Попробуйте снова.")

# Запуск асинхронного кода
if __name__ == "__main__":
    asyncio.run(main())
