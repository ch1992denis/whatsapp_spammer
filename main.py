import pywhatkit
import openpyxl
import time
import pyautogui
from pynput.keyboard import Key, Controller

keyboard = Controller()

# считываем столбец с номерами телефонов из excel-файла и заносим номера в список
book = openpyxl.open("Yutong.xlsx", read_only=True)
sheet = book.active
numbers_list = [f'+7{sheet[row][10].value.rstrip()[-10:]}' for row in range(2, sheet.max_row)]

# функция проверки номера на ошибки
def verify_number(number):
    if number[1:].isdigit() and number[2] == '9' and len(number) == 12:
        return True
    return False

# функция отправки сообщения по номеру
def send_message(number):
    try:
        pywhatkit.sendwhatmsg_instantly(
            phone_no=number,
            message= 'Добрый день! Как дела эксплуатацией и обслуживанием автобусов Ютонг? ' \
           'Есть ли нерешенные проблемы и требуется ли какая-либо помощь со стороны завода?',
            tab_close=True
        )
        time.sleep(10)
        pyautogui.click()
        time.sleep(2)
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
        print("Message sent!")
    except Exception as e:
        print(f'Message can not be delivered. {str(e)}')

# основная функция, которая в цикле проходится по номерам из списка, проверяет номер на ошибки и с паузой отправляет сообщения
def main():
    for num in numbers_list:
        if verify_number(num):
            send_message(num)
            time.sleep(15)

if __name__ == '__main__':
    main()