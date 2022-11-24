"""Created by AuF"""
import requests
import openpyxl
from tkinter.filedialog import askopenfilename
from datetime import datetime
from utils.config import LOGIN, PASSWORD

# Сайт Nikit'ы куда отправляем POST запросы
# Nikit's website where we send POST requests
url = 'https://smspro.nikita.kg/api/message'


def lead_time(func):
    """Task execution time"""
    def stopwatch():

        start_time = datetime.now()
        func()
        end_time = datetime.now()
        print(f"[*] Program has been completed for {end_time - start_time}")

    return stopwatch()


def loaner_handler() -> dict:
    """Обрабатываем клиентов у которых скоро погашение
    We process clients whose repayment is coming soon"""
    cell = 2
    client_list = {}
    wb = openpyxl.load_workbook(askopenfilename())
    sheet = wb.active

    while True:
        active_cell = sheet[f'a{cell}'].value
        if active_cell is None:
            return client_list
        else:
            sum_value = format(
                float(str(sheet[f's{cell}'].value).replace(',', '.'))
                -
                float(str(sheet[f'n{cell}'].value).replace(',', '.'))
                +
                float(str(sheet[f'k{cell}'].value).replace(',', '.')),
                '.2f')

            phone = sheet[f'u{cell}'].value
            phone = phone.split(' ')
            phone = ''.join(phone)

            date = sheet[f'p{cell}'].value
            date = date.strftime("%d/%m/%Y")

            client_list[sheet[f'b{cell}'].value] = {
                'full_name': sheet[f'a{cell}'].value,
                'date_of_m': date,
                'sum': str(sum_value),
                'delay': int(sheet[f'g{cell}'].value),
                'phone': phone
            }
            cell += 1


def create_xml(login: str, password: str) -> list:
    """Возвращает список с xml запросами
    Returns a list with xml queries"""
    list_of_xml = []
    list_of_clients = loaner_handler()

    # Перебор клиентов и подготовка тела xml документа
    # Iterating over clients and preparing the body of the xml document
    for req in list_of_clients:
        if list_of_clients[req]['delay'] > 0:
            # Сообщение для тех, у кого просрочки
            # Message for those with delays
            text = f'U vas {list_of_clients[req]["delay"]} dney prosrochki po KD #{req}, na ' \
                   f'{datetime.today().strftime("%d/%m/%Y")} summa k pogasheniu - ' \
                   f'{list_of_clients[req]["sum"]} som'
        else:
            # Сообщение для тех у кого нет просрочек
            # Message for those who do not have delays
            text = f'Напоминаем, {list_of_clients[req]["date_of_m"]} погашение по КД №{req} - ' \
                   f'{list_of_clients[req]["sum"]} сом'

        # ID запроса (может состоять из 12 знаков)
        # Request ID (can be up to 12 characters)
        send_id = f'KD{req[-6:]}D{datetime.today().strftime("%m%d")}'

        # Тело xml документа
        # The body of the xml document
        xml_body = '<?xml version="1.0" encoding="UTF-8"?>' +\
                   '<message>' +\
                   '<login>' + login + '</login>' +\
                   '<pwd>' + password + '</pwd>' +\
                   '<id>' + send_id + '</id>' +\
                   '<sender>' + 'EletKapital' + '</sender>' +\
                   '<text>' + text + '</text>' +\
                   '<phones>' +\
                   '<phone>' + list_of_clients[req]['phone'][1:] + '</phone>' +\
                   '</phones>' +\
                   '<test>' + '' + '</test>' +\
                   '</message>'
        list_of_xml.append(xml_body)
    return list_of_xml


if __name__ == '__main__':
    @lead_time
    def main():
        """===!!!RUN!!!==="""
        xml_urls = create_xml(login=LOGIN, password=PASSWORD)
        for xml in xml_urls:
            print(xml)
            xml = xml.encode('utf-8')
            r = requests.post(url, xml)
            print(r.text)
