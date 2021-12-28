import smtplib
import ssl
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import basename

import datetime
import requests
import xlwt
from bs4 import BeautifulSoup

# Замените данные для авторизации почты
YOUR_EMAIL_PASSWORD = 'YOUR_PASSWORD'
YOUR_EMAIL_ADDRESS = 'greenatom@rasskazchikov.ru'
YOUR_SMTP_SERVER = 'smtp.yandex.ru'

# Замените данный адрес на адрес получателя письма
RECIEVER_EMAIL = 'greenatom@rasskazchikov.ru'


class BsAgent:
    def __init__(self):
        pass

    def work(self):
        dollar_course_data = self.get_indicative_courses('USD/RUB')
        euro_course_data = self.get_indicative_courses('EUR/RUB')

        headers = ('Дата $', 'Курс $', 'Изменение $', 'Дата €', 'Курс €', 'Изменение €', 'Отношение курсов €/$')
        data = tuple(map(lambda x: (x[0] + x[1] + (x[1][1] / x[0][1],)), zip(dollar_course_data, euro_course_data)))
        data_with_headers = headers, *data

        rows_n = self.save_data_in_excel(data_with_headers)

        word = self.change_word_form(rows_n)

        self.send_file_via_email(
            receiver_email=RECIEVER_EMAIL,
            subject='Автоматический отчёт',
            message=f'Высылаю таблицу с курсами за прошедший месяц ({rows_n} {word})',
            path_to_file='Отчёт.xls'
        )

    def change_word_form(self, n, word_forms=('строк', 'строка', 'строки')):
        dec_rows_n = n % 10
        if n == 0 or dec_rows_n == 0 or dec_rows_n >= 5 or n in range(11, 19):
            st = word_forms[0]
        elif dec_rows_n == 1:
            st = word_forms[1]
        else:
            st = word_forms[2]
        return st

    def get_indicative_courses(self, currency='USD/RUB') -> tuple:
        end_date = datetime.datetime.now()
        year, month_i = end_date.year, end_date.month

        start_date = datetime.datetime(year, month_i, 1) - datetime.timedelta(1)

        # Set the parameters for the request
        payload = {
            'moment_start': start_date.strftime('%Y-%m-%d'),
            'moment_end': end_date.strftime('%Y-%m-%d'),
            'currency': currency,
            'language': 'ru'
        }
        xml_raw = requests.get('https://www.moex.com/export/derivatives/currency-rate.aspx', params=payload).text

        parsed_xml = BeautifulSoup(xml_raw, features="xml")

        # Select necessary data from parsed file
        data = tuple(map(lambda x: (x['moment'], float(x['value'])), reversed(parsed_xml.find_all('rate'))))

        # Add 'changes' field to the result
        result = tuple((date, value, (round(value - data[i - 1][1], 4)) if i > 0 else None)
                       for i, (date, value) in enumerate(data))

        # Get rid of data from last month
        this_month_data = tuple(filter(lambda x: x[0].startswith(end_date.strftime('%Y-%m')), result))

        return this_month_data

    def send_file_via_email(self, receiver_email, subject, message, path_to_file=None):

        msg = MIMEMultipart()
        msg['From'] = YOUR_EMAIL_ADDRESS
        msg['To'] = receiver_email
        msg['Subject'] = subject

        msg.attach(MIMEText(message))

        if path_to_file is not None:
            with open(path_to_file, "rb") as fil:
                part = MIMEApplication(
                    fil.read(),
                    Name=basename(path_to_file)
                )
            part['Content-Disposition'] = 'attachment; filename="%s"' % basename(path_to_file)
            msg.attach(part)

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(YOUR_SMTP_SERVER, 465, context=context) as server:
            server.login(YOUR_EMAIL_ADDRESS, YOUR_EMAIL_PASSWORD)
            server.sendmail(YOUR_EMAIL_ADDRESS, receiver_email, msg.as_string())
        pass

    def save_data_in_excel(self, data) -> int:
        FINANCE_STYLE = xlwt.XFStyle()
        FINANCE_STYLE.num_format_str = '#,##0.00 ₽'

        wb = xlwt.Workbook()
        ws = wb.add_sheet('Отчёт')

        for row_i, data_row in enumerate(data):
            for column_i, cell_data in enumerate(data_row):

                cwidth = ws.col(column_i).width
                if (len(str(cell_data)) * 367) > cwidth:
                    ws.col(column_i).width = (len(cell_data) * 367)

                if column_i in (1, 2, 4, 5):
                    ws.write(row_i, column_i, cell_data, FINANCE_STYLE)
                else:
                    ws.write(row_i, column_i, cell_data, xlwt.easyxf("align: horiz center"))

        wb.save('Отчёт.xls')

        return len(ws._Worksheet__rows) - 1


if __name__ == '__main__':
    agent = BsAgent()
    agent.work()
