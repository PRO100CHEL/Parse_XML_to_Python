import pandas as pd
import xml.etree.ElementTree as ET


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
    with pd.ExcelWriter(excel_file) as writer:
        result_df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Автоширина столбцов и форматирование чисел
        for column in result_df.columns:
            column_width = max(result_df[column].astype(str).map(len).max(), len(column))
            col_idx = result_df.columns.get_loc(column)
            writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width, None, {'num_format': '#,##0.00'})


usd_xml_file = 'USD_RUB.xml'
jpy_xml_file = 'JPY_RUB.xml'
excel_file = 'Полученный файл.xlsx'
xml_to_excel(usd_xml_file, jpy_xml_file, excel_file)
