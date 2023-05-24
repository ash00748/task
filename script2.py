import sys
import pandas as pd
import xml.etree.ElementTree as ET


def excel_to_xml(file: str, output_file: str):
    """
    функция, генерирующая xml файл на основе excel спредшита
    Parameters
    ----------
    file : str
        имя файла xlsx
    output_file : str
        имя файла xml, в который будет записан сгенерированный результат
    """
    # Загружаем spreadsheet в объект pandas
    # dataframe с основными даннами для работы
    try:
        df = pd.read_excel(
            file,
            sheet_name="Sheet1",
            header=4,
            names=[
                "CERTNO",
                "CERTDATE",
                "STATUS",
                "IEC",
                "EXPNAME",
                "BILLID",
                "SDATE",
                "SCC",
                "SVALUE",
            ],
        )

    # dataframe - первые три строки с важной информацией (шапка)
        head = pd.read_excel(
            file,
            nrows=3,
            names=["name", "CERTDATA"],
            header=None,
            index_col=0,
            usecols="A:B",
        )
    except FileNotFoundError:
        print(f'Excel спредшит \"{file}\" не найден')
        sys.exit(1)

    # код валюты (доллары США)
    code_usa = 'R01235'

    root = ET.Element("CERTDATA")  # задаем root элемент
    tree = ET.ElementTree(root)

    filename = ET.SubElement(root, "FILENAME")  # задаем элемент filename
    filename.text = str(head.at["FileName", "CERTDATA"])

    envelope = ET.SubElement(root, "ENVELOPE")  # задаем элемент envelope
    
    # словарь с курсом доллара
    dollar_values = {} 
    for row in df.index:  # для каждой строки данных задаем элемент ecert
        ecert = ET.SubElement(envelope, "ECERT")
        for column in df.columns:  # для каждого поля данных создаем соотвествующий элемент
            entry = ET.SubElement(ecert, column)
            data = df.at[row, column]
            if column == "CERTDATE" or column == "SDATE":
                data = data.date()  # необходимое форматирование
            if column == "IEC":
                data = f"{data:010d}"
            if column == "EXPNAME":
                data = f"\"{data}\""
            if column == "SVALUE" and data % 10 == 0:
                data = int(data)
            entry.text = str(data)

        # переформатируем дату в формат, который используется в api cbr
        date = (df.at[row, 'SDATE']).date()
        date = f'{date.day:02}/{date.month:02}/{date.year}'
        svalue = df.at[row, 'SVALUE']

        # если курс для этой даты ещё не загружали
        if date not in dollar_values.keys(): 
            url = f"http://www.cbr.ru/scripts/XML_daily.asp?date_req={date}"
            xml_cbr = pd.read_xml(url, encoding='cp1251')

            # индекс строки, соответствующей нужной валюте
            index_usa = ((xml_cbr.index[xml_cbr['ID'] == code_usa]).to_list())[0]  
            # курс нужной валюты лежит в формате str 123,123
            value_usa = float((xml_cbr.at[index_usa, 'Value']).replace(',', '.'))
            dollar_values[date] = value_usa
        else: # если курс для даты уже загружен
            value_usa = dollar_values[date]
        # вычисляем значение атрибута SVALUEUSD и округляем согласно требованиям
        svalueus = round(svalue/value_usa, 2)
        # добавляем новый атрибут в xml дерево
        entry = ET.SubElement(ecert, 'SVALUEUSD')
        entry.text = str(svalueus)

    # красиво записываем в файл
    ET.indent(tree, space="\t", level=0)
    tree.write(f'{output_file}', encoding="UTF-8", xml_declaration=True)


if __name__ == "__main__":
    file_name = sys.argv[1]
    out_file = sys.argv[2]
    excel_to_xml(file_name, out_file)