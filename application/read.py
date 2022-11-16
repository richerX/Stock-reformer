from application.constants import *


# Чтение остатков
def read_storage_openpyxl(page, products):
    # Чтение магазинов, которые есть в остатках
    shops = []
    for i in range(10):
        letter = letters[i + 2]
        value = page[letter + '8'].value
        if value in storage_shops_update.keys():
            shops.append(products[storage_shops_update[value]])
        elif value == "Итог":
            shops.append(products["All"])
        else:
            shops.append(dict())

    # Чтение остатков
    current_row, article = 11, None
    while page["B" + str(current_row)].value not in ["Итог", None]:
        article_split = str(page["B" + str(current_row)].value).split()

        # Цветная клетка - клетка с артикулом
        if page["B" + str(current_row)].fill.fgColor.rgb != "FFFFFFFF":
            brand_index = next((i for i in range(len(article_split)) if article_split[i] in brand_names), 2)
            article = " ".join(article_split[:brand_index]).strip()
            articles_in_storage.add(article)

        # Не цветная клетка - клетка с размерами
        else:
            size = article_split[0].rstrip(",").lstrip("0")
            size = size if size != "ONE" else "ONE SIZE"
            # Добавление значений ряда
            current_key = (article, size)
            for i in range(len(shops)):
                current_cell = page[letters[i + 2] + str(current_row)]
                # Считывания значения клетки
                if current_cell.value is not None:
                    if shops[i].get(current_key) is None:
                        shops[i][current_key] = 0
                    shops[i][current_key] += int(current_cell.value)

        # Переход на следующий ряд
        current_row += 1
    return current_row
