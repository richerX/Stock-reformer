import traceback
import openpyxl
import copy

from application.read import *
from application.write import *
from application.print import *


# Получить строку с разницей во времени
def get_time_string(time_stamp):
    return f"{str(round(time.time() - time_stamp, 3))} сек."


# Получение списка рядов артикулов
def get_article_rows(page):
    article_rows_answer = []
    current_row, current_slips = 2, 0
    while current_slips < max_article_step:
        name = type(page[f"B{current_row}"]).__name__
        if name == 'MergedCell':
            article_rows_answer.append(current_row)
            current_row += min_article_step
            current_slips = 0
        else:
            current_row += 1
            current_slips += 1
    return article_rows_answer


# Вывод остатков в консоль
def print_products(products):
    if print_products_mode:
        for key, value in products.items():
            print(key)
            for product in value:
                print(f"\t{product}")


# Основной метод программы
def main():
    main_begin_time = time.time()
    create_folder(text_files_folder, True)
    create_folder(table_files_folder, False)
    write_in_file(log_filepath, f"Wardrobe (version {build_version})")

    # Чтение остатков
    success_read_storage = False
    products = copy.deepcopy(new_content)
    try:
        wb = openpyxl.load_workbook(f"{table_files_folder}/{storage_filepath}")
        console_print(storage_filepath, "Открытие таблицы", "", get_time_string(main_begin_time))
        storage_begin_time = time.time()
        position = read_storage_openpyxl(wb.active, products)
        console_print(wb.active.title, str(position), get_time_string(storage_begin_time))
        success_read_storage = True
    except Exception as exception:
        error_log(f"Ошибка при чтении остатков: exception = {exception}")
    print_products(products)

    # Заполнение таблицы
    current_tables = existing_tables if success_read_storage else list()
    for table_filename in current_tables:
        table_begin_time = time.time()
        try:
            wb = openpyxl.load_workbook(f"{table_files_folder}/{table_filename}")
        except Exception as exception:
            error_log(f"Ошибка при открытии таблицы: table = {table_filename}, exception = {exception}")
            continue
        table_type = table_filepaths[table_filename][1]
        columns = columns_all[table_type]
        keys = keys_all[table_type]
        file_content = copy.deepcopy(new_content)
        console_print(table_filename, "Открытие таблицы", "", get_time_string(table_begin_time))

        # Заполнение страницы
        for page_number in range(len(wb.worksheets)):
            if wb.worksheets[page_number].title in skip_lists:
                continue
            page_begin_time = time.time()
            page = wb.worksheets[page_number]
            article_rows = get_article_rows(page)
            page_content = copy.deepcopy(new_content)
            for row in range(len(article_rows)):
                step = article_rows[row + 1] - article_rows[row] if row < len(article_rows) - 1 else max_article_step
                for current_columns in columns:
                    try:
                        fill_table_product_cell(page, article_rows[row], step, current_columns, keys, products, page_content, file_content)
                    except Exception as exception:
                        print(page, article_rows[row], step, current_columns)
                        error_log(f"Ошибка при заполнении таблицы: table = {table_filename}, page = {page.title}, row = {article_rows[row]}, step = {step}, current_columns = {current_columns}, exception = {exception}")

            # Вывод блока справа-сверху
            try:
                headers = headers_all[table_type]
                diapason = diapason_all[table_type]
                fill_info_block(page, page_content, headers, diapason)
                if page_number == len(wb.worksheets) - 1:
                    headers = headers_all_lp[table_type]
                    diapason = diapason_all_lp[table_type]
                    fill_info_block(page, file_content, headers, diapason)
            except Exception as exception:
                error_log(f"Ошибка при заполнении блока справа-сверху: table = {table_filename}, page = {page.title}, exception = {exception}")

            # Вывод в консоль
            last_article_row = article_rows[-1] if len(article_rows) > 0 else 0
            console_print(page.title, str(last_article_row), get_time_string(page_begin_time), update_console = False)

        # Сохранение таблицы
        try:
            wb.save(f"{table_files_folder}/{table_filename}")
        except Exception as exception:
            error_log(f"Ошибка при сохранении таблицы: table = {table_filename}, exception = {exception}")

    # Вывод в конце программы
    write_in_file(articles_not_in_tables_filepath, "Aртикулы, которых нет в таблицах:", list(articles_in_storage - articles_in_tables))
    write_in_file(articles_not_in_storage_filepath, "Aртикулы, которых нет в остатках:", list(articles_in_tables - articles_in_storage))
    console_print("Итого", "Общее время работы", "", get_time_string(main_begin_time))


if __name__ == "__main__":
    try:
        main()
    except Exception as critical_exception:
        traceback.print_exception(*sys.exc_info())
        input()
        error_log(f"Критическая ошибка выполнения программы: exception = {critical_exception}")
        with open(text_files_folder + "/" + error_log_filepath + ".txt", "a", encoding = "utf-8") as file:
            traceback.print_exception(*sys.exc_info(), file = file)
            file.write("\n")
    input("\nДля завершения нажмите Enter...")
