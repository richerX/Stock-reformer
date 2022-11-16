import os
import sys
import time
from rich.columns import Columns
from rich.panel import Panel

from application.constants import *


# Создать папку
def create_folder(folder, clean_folder = False):
    try:
        os.mkdir(folder)
    except FileExistsError:
        if clean_folder:
            filenames = [file for file in os.listdir(folder) if file.endswith(".txt")]
            for filename in filenames:
                file = open(text_files_folder + "/" + filename, "w", encoding = "utf-8")
                file.close()


# Вывод в консоль
def console_print(*words, update_console = True):
    console_log(words)
    table.add_row(*prettify_words(words))
    if update_console and not print_products_mode:
        os.system('CLS')
        console.print(table)


# Корректировка под python rich
def prettify_words(words):
    words = list(map(lambda x: str(x), words))
    if len(words) != 3 and len(table.rows) > 0:
        table.rows[-1].end_section = True
    if len(words) == 1 and len(table.rows) > 0:
        table.columns[-1].no_wrap = True
    if len(words) == 4:
        words[0] = Columns([Panel(f"[{color_filename}]" + words[0])])
        words[3] = f"[{color_time}]" + words[3]
    if len(words) == 3:
        words = [""] + words
        words[3] = f"[{color_time}]" + words[3]
    if len(words) == 2:
        words[0] = f"[{color_special}]" + words[0]
        words[1] = f"[{color_special}]" + words[1]
    if len(words) == 1:
        words[0] = f"[{color_error}]" + words[0]
    return words


# Запись в лог
def console_log(words):
    words = [str(i) for i in words]
    if len(words) == 0:
        string1 = "|" + "-" * console_table_width * 12 + "|"
        string2 = ""
    elif len(words) == 1:
        string1 = "|" + "-" * console_table_width * 12 + "|"
        string2 = "|" + words[0].center(console_table_width * 12) + "|"
        if len(words[0]) > console_table_width * 12 - 2:
            string2 = "|" + (words[0][:console_table_width * 12 - 5] + "...").center(console_table_width * 12) + "|"
    elif len(words) == 2:
        string1 = "|" + "-" * console_table_width * 12 + "|"
        string2 = "|" + words[0].center(console_table_width * 9) + "|" + words[1].center(console_table_width * 3 - 1) + "|"
    elif len(words) == 3:
        string1 = "|" + "".center(console_table_width * 5) + "|" + "-" * (console_table_width * 7 - 1) + "|"
        string2 = "|" + "".center(console_table_width * 5) + "|" + words[0].center(console_table_width * 4 - 1) + "|" + words[1].center(console_table_width * 1 - 1) + "|" + words[2].center(console_table_width * 2 - 1) + "|"
    elif len(words) == 4:
        string1 = "|" + "-" * console_table_width * 12 + "|"
        string2 = "|" + words[0].center(console_table_width * 5) + "|" + words[1].center(console_table_width * 4 - 1) + "|" + words[2].center(console_table_width * 1 - 1) + "|" + words[3].center(console_table_width * 2 - 1) + "|"
    else:
        string1 = "|" + "-" * console_table_width * 12 + "|"
        string2 = "|", f"Неверное кол-во аргументов для функции console_print(): words = {words}".center(console_table_width), "|"
    write_in_file(log_filepath, string1)
    write_in_file(log_filepath, string2)


# Запись в лог ошибок
def error_log(text):
    console_print(text)
    write_in_file(error_log_filepath, f"[{time.strftime('%H:%M:%S', time.localtime())}] {text}\n[ERROR]{sys.exc_info()}\n")


# Запись текста в файл
def write_in_file(filename, text, array = None):
    with open(text_files_folder + "/" + filename, "a", encoding = "utf-8") as file:
        file.write(text + "\n")
        if array is not None:
            array.sort()
            file.write("\n")
            for element in array:
                file.write(str(element) + "\n")
