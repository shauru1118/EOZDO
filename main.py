PROSROCHKI_FILE_NAME = "1.xls"
VIPOLNENO_FILE_NAME = "2.xls"
RESULT_FILE_NAME = "res_prem.xlsx"

from time import sleep 
import time

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

prosrochki = exel.ExelWB(PROSROCHKI_FILE_NAME, RESULT_FILE_NAME)
prosrochki.do_prosrochki()
prosrochki.end()



vipolneno = exel.ExelWB(VIPOLNENO_FILE_NAME, RESULT_FILE_NAME)
vipolneno.do_vipolneno()
vipolneno.end()

import tqdm

print("\n\n")

for i in tqdm.tqdm(range(300), colour="white"):
    sleep(0.001)

for i in tqdm.tqdm(range(300), colour="blue"):
    sleep(0.001)


for i in tqdm.tqdm(range(300), colour="red"):
    sleep(0.001)

sleep(0.8)

from termcolor import colored

console_width = os.get_terminal_size().columns

print(f"\n\n{"=" * console_width}\n" + colored("<<<< Данные выше предназначены для отладки и контроля, они вам не нужны >>>>",
                                                                       "green"))

print(colored("\n" + "Готово!\n" + f"Результаты сохранены в файле: {RESULT_FILE_NAME}\n" + 
      "Он откроется сам, а вы дочитайте этот текст :).\n" + 
      "Проверьте файл, чтобы убедиться, что все сработало правильно :)\n" + 
      "Я не прощаюсь с Вами, а говорю Вам до новых встреч!\n", "light_green"))


end = input(f"{"=" * console_width}\n\nНажмите Enter, чтобы продолжить...")

os.startfile(RESULT_FILE_NAME)
