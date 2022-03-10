import os.path
import traceback

from time import sleep

from colorama import init, Fore, Style
import win32console
from pyfiglet import Figlet

from core import read_excel_table, process_user


def print_head() -> None:
    figlet = Figlet(font="speed")
    banner = figlet.renderText("Fast practice")
    print(f"{Style.BRIGHT}{Fore.MAGENTA}{banner}")
    print("All rights reserved by Pochatkov A.R.\n")
    print(f"{Fore.RED}telegram: @andrpocc\n" + "vk: andrpocc\n")


def input_path_to_excel() -> str:
    print(
        f"{Fore.MAGENTA}Перенесите {Fore.RED}Excel{Fore.MAGENTA} файл в окно программы и нажмите клавишу ввода:\n{Fore.RED}{Style.BRIGHT}"
    )
    path_to_excel = input().replace('"', "")
    print(f"\n{Fore.MAGENTA}PDF файлы будут сохранены по указанному выше пути!\n\n{Fore.RED}Таблица из файла:\n")
    sleep(1)
    return path_to_excel


def main():
    init()
    win32console.SetConsoleTitle("Fast Practice")
    print_head()
    excel_file_path = input_path_to_excel()
    save_path = os.path.split(excel_file_path)[0]

    table = read_excel_table(excel_file_path)
    print(table)
    print(f"{Fore.MAGENTA}\nВведите номер группы:{Fore.RED}\n")
    group = input("Введите номер группы: ")
    print()
    broken = []
    for _, row in table.iterrows():
        name = str(row[1])
        mark = str(row[2])
        context = {"name": name, "mark": mark, "group": group}
        print(f"{Fore.MAGENTA}Обрабатываю студента: {Fore.RED}{name}")
        try:
            process_user(context, save_path)
            print(f"{Fore.MAGENTA}Успешно!\n")
        except Exception:
            print()
            print(traceback.format_exc())
            print(f"{Fore.MAGENTA}Возникла ошибка!\n")
            broken.append(name)
            input("Нажмите клавишу ввода, чтобы продолжить..")
    if broken:
        print(f"Следующие студенты не были обработаны из-за ошибки:{Fore.RED}")
        for name in broken:
            print(name)
    print(f"Все записи обработаны и сохранены по пути: {Fore.RED}{save_path}\n")
    input("Нажмите клавишу ввода для выхода из программы..")


if __name__ == "__main__":
    main()
