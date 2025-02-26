PROSROCHKI_FILE_NAME = "1.xls"
VIPOLNENO_FILE_NAME = "2.xls"
RESULT_FILE_NAME = "res_prem.xlsx"

from time import sleep 

print("ВНИМАНИЕ!\n" + 
      f"Убедитесь, что два файла: {PROSROCHKI_FILE_NAME} (просрочки) и {VIPOLNENO_FILE_NAME} (закрыто) " + 
      "присутствуют на рабочем столе, откуда вы и запускаете программу, и они закрыты.\n" + 
      "В ином случае закройте программу перенесите, переименуйте и закройте все файлы и запустите программу заново.")
start = input("Нажмите Enter, чтобы продолжить...")

import os

# Проверка наличия двух необходимых файлов

if not os.path.exists(PROSROCHKI_FILE_NAME) or not os.path.exists(VIPOLNENO_FILE_NAME):
    print("Один из файлов не найден!")
    exit()

# Подключение необходимых библиотек

import exel


exel.ExelWB.create_res_premia(RESULT_FILE_NAME)

prosrochki = exel.ExelWB(PROSROCHKI_FILE_NAME, RESULT_FILE_NAME, {})
pros_svod_table = prosrochki.do_prosrochki()
prosrochki.end()

vipolneno = exel.ExelWB(VIPOLNENO_FILE_NAME, RESULT_FILE_NAME, pros_svod_table)
vipo_svod_table = vipolneno.do_vipolneno()
vipolneno.end()

svod_table = exel.ExelWB(None, RESULT_FILE_NAME, vipo_svod_table)
svod_table.do_svod()

svod_table.end()

sleep(0.5)


from termcolor import colored

console_width = os.get_terminal_size().columns

print(f"\n\n{"=" * console_width}")

print(colored("\n" + "Готово!\n" + f"Результаты сохранены в файле: {RESULT_FILE_NAME}\n" + 
      "Он откроется сам, а вы дочитайте этот текст :).\n" + 
      "Проверьте файл, чтобы убедиться, что все сработало правильно :)\n" + 
      "Я не прощаюсь с Вами, а говорю Вам до новых встреч!\n", "light_green"))


end = input(f"{"=" * console_width}\n\nНажмите Enter, чтобы продолжить...")

os.startfile(RESULT_FILE_NAME)
