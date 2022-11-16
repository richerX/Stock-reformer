from application.constants import *
from openpyxl.utils.cell import *


# Заполнение всей ячейки только одного товара
def fill_table_product_cell(page, row, step, columns, keys, products, page_content, file_content):
    if page[columns[0] + str(row)].value is None:
        return
    article = str(page[columns[0] + str(row)].value).strip()
    articles_in_tables.add(article)

    # Заполнение ряда
    for current_row in range(row + 2, row + step):
        # Определение размера
        size = page[columns[0] + str(current_row)].value
        size = maximo_size_update[size] if size in maximo_size_update.keys() else size  # maximo
        size = str(size).replace(".", ",") if type(size) == float else str(size)  # pretty ballerinas and maximo
        page_content["Sizes"].add(size)
        file_content["Sizes"].add(size)

        # Заполнение клетки
        for index in range(len(columns) - 1):
            column = columns[index + 1]
            shop_products = products[keys[index]]
            current_content_page = page_content[keys[index]]
            current_content_book = file_content[keys[index]]

            page[column + str(current_row)] = None
            if shop_products.get((article, size)) is not None:
                if current_content_page.get(size) is None:
                    current_content_page[size] = 0
                if current_content_book.get(size) is None:
                    current_content_book[size] = 0
                quantity = shop_products[(article, size)]
                current_content_page[size] += quantity
                current_content_book[size] += quantity
                page[column + str(current_row)] = quantity


# Заполнение блока справа-сверху
def fill_info_block(page, content, headers, diapason):
    # Очистка места для блока справа-сверху
    first_column = letters[letters.index(diapason[0]) - 1]
    last_column = letters[letters.index(diapason[-1]) + 10]
    for row in page[first_column + "1" + ":" + last_column + "100"]:
        for cell in row:
            coords = coordinate_from_string(cell.coordinate)
            down_cell = page.cell(coords[1] + 1, column_index_from_string(coords[0]))
            right_cell = page.cell(coords[1], column_index_from_string(coords[0]) + 1)
            # Вылетает ошибка при попытке обнулить MergedCell
            # Левый верхний угол MergedCell не считается за MergedCell
            if not (type(cell).__name__ == 'MergedCell' or (type(down_cell).__name__ == 'MergedCell' and type(right_cell).__name__ == 'MergedCell')):
                cell.value = None
                cell.style = "Normal"

    # Получение всех размеров таблицы
    all_sizes, numeric_sizes, string_sizes = list(), list(), list()
    maximo_enabled = any([size in maximo_size_update.values() for size in content["Sizes"]])
    content_sizes = [size for size in content["Sizes"] if size != 'None' and size not in maximo_size_update.values()]
    for current_size in content_sizes:
        try:
            numeric_sizes.append(int(current_size))
        except ValueError:
            string_sizes.append(current_size)
    for array_sizes in [sorted(numeric_sizes), sorted(string_sizes)]:
        all_sizes += [str(element) for element in array_sizes]
    if maximo_enabled:
        all_sizes += list(maximo_size_update.values())

    # Стилизация блока справа-сверху
    for index in range(len(diapason)):
        page[diapason[index] + str(height_of_right_block)] = headers[index]
        page[diapason[index] + str(height_of_right_block)].alignment = Alignment(horizontal="center", vertical="center")
        page.column_dimensions[diapason[index]].width = 12
    make_border(page, height_of_right_block, len(all_sizes) + height_of_right_block, diapason[0], diapason[-1])

    # Заполнение блока справа-сверху
    for index in range(len(all_sizes)):
        current_row = index + height_of_right_block + 1
        current_size = all_sizes[index]
        if current_size in maximo_size_update.values():
            page[diapason[0] + str(current_row)] = " " + current_size
        else:
            page[diapason[0] + str(current_row)] = current_size
        page[diapason[0] + str(current_row)].alignment = Alignment(horizontal = "center", vertical = "center")

        # Заполнение ряда
        total_in_row = 0
        for j in range(len(diapason) - 2):
            current_column = diapason[j + 1]
            current_store = page[current_column + str(height_of_right_block)].value
            current_content = content[headers_to_keys[current_store]]
            if current_content.get(current_size) is None:
                page[current_column + str(current_row)] = 0
            else:
                page[current_column + str(current_row)] = current_content[current_size]
                total_in_row += current_content[current_size]
            page[current_column + str(current_row)].alignment = Alignment(horizontal = "center", vertical = "center")
        page[diapason[-1] + str(current_row)] = total_in_row
        page[diapason[-1] + str(current_row)].alignment = Alignment(horizontal = "center", vertical = "center")


# Создать обводку для блока справа-сверху
def make_border(page, first_row, last_row, first_column, last_column):
    index_of_first_column = letters.index(first_column)
    index_of_last_column = letters.index(last_column)

    # Верхняя и нижняя граница
    for column in letters[index_of_first_column:index_of_last_column + 1]:
        page[column + str(first_row)].border = top_border
        page[column + str(last_row)].border = bottom_border

    # Правая и левая границы
    for local_row in range(first_row, last_row + 1):
        page[first_column + str(local_row)].border = left_border
        page[last_column + str(local_row)].border = right_border

    # Угловые границы
    page[first_column + str(first_row)].border = left_top_border
    page[first_column + str(last_row)].border = left_bottom_border
    page[last_column + str(first_row)].border = right_top_border
    page[last_column + str(last_row)].border = right_bottom_border
