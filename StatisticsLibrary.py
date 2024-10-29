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


# Метод для добавления количества пропусков студенту
async def add_absences(student_name, absences, worksheet):
    try:
        student_cell = await worksheet.find(student_name)
        if not student_cell:
            print(f"Студент '{student_name}' не найден.")
            return

        # Предполагается, что "Дата недели" - третий столбец
        absence_column = student_cell.col + 2  # 3-й столбец для "с 21 по 27 октября"
        total_absences_column = student_cell.col + 3  # 4-й столбец для "За полугодие"

        # Используем await для получения значения из ячейки
        current_absences_cell = await worksheet.cell(student_cell.row, absence_column)
        current_absences = current_absences_cell.value  # Получаем значение ячейки

        if current_absences:  # Если у студента уже есть пропуски за этот период
            current_absences = int(current_absences)
        else:
            current_absences = 0

        # Добавляем пропуски
        new_total_absences = current_absences + absences
        await worksheet.update_cell(student_cell.row, absence_column, new_total_absences)
        print(
            f"Добавлено {absences} пропусков для студента {student_name}. Всего пропусков за неделю: {new_total_absences}.")

        # Обновляем колонку "За полугодие"
        total_absences_cell = await worksheet.cell(student_cell.row, total_absences_column)
        total_absences = total_absences_cell.value  # Получаем значение ячейки "За полугодие"

        if total_absences:  # Если есть уже сохраненные пропуски за полугодие
            total_absences = int(total_absences) + absences
        else:
            total_absences = absences  # Если пропусков не было, устанавливаем новые

        await worksheet.update_cell(student_cell.row, total_absences_column, total_absences)
        print(f"Обновлено общее количество пропусков за полугодие для студента {student_name}: {total_absences}.")

    except Exception as e:
        print(f"Ошибка при добавлении пропусков: {e}")


# Метод для обновления заголовка даты
async def update_date_column_header(new_date_range, worksheet):
    try:
        absence_column = 3  # 3-й столбец для "с 21 по 27 октября"

        # Очищаем все ячейки в столбце с пропусками
        cells = await worksheet.get_all_values()
        for i in range(1, len(cells)):  # Пропускаем заголовок
            await worksheet.update_cell(i + 1, absence_column, '')  # Очищаем ячейки

        await worksheet.update_cell(1, absence_column, new_date_range)  # Обновляем заголовок третьего столбца
        print(f"Заголовок даты обновлен на: {new_date_range} и все пропуски очищены.")
    except Exception as e:
        print(f"Ошибка при обновлении заголовка: {e}")


# Метод для вывода полной статистики
async def print_full_statistics(worksheet):
    try:
        # Считываем все данные из таблицы
        records = await worksheet.get_all_values()

        # Определяем ширину колонок для красивого вывода
        col_widths = [max(len(str(cell)) for cell in col) for col in zip(*records)]

        # Форматированный вывод заголовка
        header = records[0]
        header_format = " | ".join(f"{header[i]:<{col_widths[i]}}" for i in range(len(header)))
        print(header_format)
        print("-" * len(header_format))

        # Вывод данных построчно
        for row in records[1:]:
            row_format = " | ".join(f"{row[i]:<{col_widths[i]}}" for i in range(len(row)))
            print(row_format)

    except Exception as e:
        print(f"Ошибка при выводе статистики: {e}")


# Основная программа для тестирования методов
async def main():
    client = await agcm.authorize()
    sheet = await client.open_by_key('1yB02wlZ6p12oM7hOvPAknFyC-aSUSaAoBJhB0Ru3VhE')
    worksheet = await sheet.get_worksheet(0)  # Работает с первым листом

    print("Выберите действие:")
    print("1. Поменять дату")
    print("2. Добавить пропуски")
    print("3. Вывести полную статистику")

    choice = input("Введите номер действия (1, 2 или 3): ").strip()

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
    else:
        print("Неверный выбор. Попробуйте снова.")


# Запуск асинхронного кода
if __name__ == "__main__":
    asyncio.run(main())
