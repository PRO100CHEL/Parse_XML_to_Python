import pandas as pd
import xml.etree.ElementTree as ET
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os


def parse_xml(xml_file, secid_filter):
    # Парсинг XML файла
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Инициализация списка для хранения данных
    data = []

    # Извлечение данных из строк (rows)
    for row in root.find('.//data[@id="securities"]/rows').findall('row'):
        tradedate = row.attrib.get('tradedate')
        tradetime = row.attrib.get('tradetime')
        secid = row.attrib.get('secid')
        rate = float(row.attrib.get('rate')) if row.attrib.get('rate') else None
        clearing = row.attrib.get('clearing')

        # Основной клиринг
        if clearing == 'vk' and secid == secid_filter:
            data.append([tradedate, rate, tradetime])

    return data


def xml_to_excel(usd_xml_file, jpy_xml_file, excel_file):
    # Парсинг данных из XML файлов
    usd_data = parse_xml(usd_xml_file, "USD/RUB")
    jpy_data = parse_xml(jpy_xml_file, "JPY/RUB")

    # Преобразование данных в DataFrame
    usd_df = pd.DataFrame(usd_data, columns=["Дата USD/RUB", "Курс USD/RUB", "Время USD/RUB"])
    jpy_df = pd.DataFrame(jpy_data, columns=["Дата JPY/RUB", "Курс JPY/RUB", "Время JPY/RUB"])

    # Объединение данных по дате и времени
    result_df = pd.merge(usd_df, jpy_df, left_on=["Дата USD/RUB", "Время USD/RUB"],
                         right_on=["Дата JPY/RUB", "Время JPY/RUB"])

    # Расчет результата
    result_df["Результат"] = result_df["Курс USD/RUB"] / result_df["Курс JPY/RUB"]

    # Сохранение DataFrame в Excel файл
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Автоширина столбцов и форматирование чисел
        for column in result_df.columns:
            column_width = max(result_df[column].astype(str).map(len).max(), len(column))
            col_idx = result_df.columns.get_loc(column)
            writer.sheets['Sheet1'].column_dimensions[
                writer.sheets['Sheet1'].cell(1, col_idx + 1).column_letter].width = column_width
            # Финансовый формат
            for cell in writer.sheets['Sheet1'].iter_rows(min_row=2, min_col=col_idx + 1, max_col=col_idx + 1,
                                                          max_row=len(result_df) + 1):
                cell[0].number_format = '#,##0.00'

    return len(result_df)


def get_correct_form(number):
    if 11 <= number % 100 <= 19:
        return f"{number} строк"
    elif number % 10 == 1:
        return f"{number} строка"
    elif 2 <= number % 10 <= 4:
        return f"{number} строки"
    else:
        return f"{number} строк"


def send_email(excel_file, recipient_email, num_rows):
    # Параметры отправки письма
    sender_email = "4stonishing_m@vk.com"  # Ваш email
    sender_password = "12345"  # Ваш пароль

    subject = "Поддержка RPA"
    body = f"Здравствуйте!\n\nВо вложении отчет по курсам валют. В файле содержится {get_correct_form(num_rows)}.\n\nhttps://github.com/PRO100CHEL/Parse_XML_to_Python\n\nС уважением,\nМильдзихов А.Д."

    # Создание письма
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    # Вложение файла
    attachment = open(excel_file, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(excel_file)}")

    msg.attach(part)
    attachment.close()

    # Настройка SMTP сервера и отправка письма
    with smtplib.SMTP('mail.ru', 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())


usd_xml_file = 'USD_RUB.xml'
jpy_xml_file = 'JPY_RUB.xml'
excel_file = 'Полученный файл.xlsx'
num_rows = xml_to_excel(usd_xml_file, jpy_xml_file, excel_file)
send_email(excel_file, '4stonishing_m@vk.com', num_rows)
