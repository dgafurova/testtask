import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime as dt


def task1(inputFileName = 'test_input.xlsx', outputFileName='output_xml.xml'):
    '''
    The task1 function opens an xlsx file and generates an xml file of the given
    structure with the information from an xlsx file.

    Parameters:
    -----------

    inputFileName - the name of the input xlsx file, default values is test_input.xlsx, str

    outputFileName - the name of the output xml file, default values is output_xml.xml, str

    '''
    # Чтение xlsx файла
    # Reading xlsx file
    data = pd.read_excel('test_input.xlsx', header=None, dtype=str)

    # Считывание информации из первых трех строк в отдельные переменные
    # Reading information from the first three lines into separate variables
    proccessingDateText, fileNumberText, fileNameText = data.iloc[:3, 1]

    # Удаление первых четырех строчек таблицы
    # Deleting the first four lines of the table
    data.columns = list(data.iloc[4])
    data = data.iloc[5:, :].reset_index(drop=True)

    # Предобработка данных
    # Data preprocessing

    # Добавление кавычек значениям столбца Client
    # Adding the quotation marks to the values in the Client column
    data['Client'] = data['Client'].map(lambda x: "\"" + x + "\"")

    # Перевод дат в нужный формат
    # Converting dates to the desired format
    data['Issuance Date'] = pd.to_datetime(data['Issuance Date']).dt.strftime('%Y-%m-%d')
    data['SB Date'] = pd.to_datetime(data['SB Date']).dt.strftime('%Y-%m-%d')

    # Перевод IE Code в формат с 10 цифрами, как в Excel файле
    # Converting IE Code to a 10-numbers format, as in an Excel file
    data['IE Code'] = data['IE Code'].str.zfill(10)

    # Словарь соответствий названий атрибутов в xml файле и xlsx файле (если в будущем названия атрибутов поменяются
    # или нужно будет работать с новыми атрибутами, эти изменения можно учесть в словаре)
    # Dictionary of matches of attribute names in an xml file and in an xlsx file (if in the future the attribute names change
    # or it will be neccesary to work with new attributes, these changes can be taken into account in this dictionary)
    xml2excelDict = {'CERTNO': 'Ref no',
                     'CERTDATE': 'Issuance Date',
                     'STATUS': 'Status',
                     'IEC': 'IE Code',
                     'EXPNAME': 'Client',
                     'BILLID': 'Bill Ref no',
                     'SDATE': 'SB Date',
                     'SCC': 'SB Currency',
                     'SVALUE': 'SB Amount'}

    # Создание элементов xml файла
    # Creating elements for an xml file

    # Создание атрибута CERTDATA
    # Creating a CERTDATA atribute
    root = ET.Element('CERTDATA')

    # Создание атрибута FILENAME и установка его значения данными из исходного xlsx файла
    # Creating a FILENAME atribute and setting its values from xlsx file
    filename = ET.SubElement(root, 'FILENAME')
    filename.text = fileNameText

    # Создание атрибута ENVELOPE
    # Crerating an ENVELOPE atribute
    envelope = ET.SubElement(root, 'ENVELOPE')

    # Цикл, в котором данные из xlsx файла добавляются в атрибуты xml файла
    # A loop in which data from an xlsx file is added to the attributes of an xml file
    for iter, row  in data.iterrows():
        # Создание атрибута ECERT
        # Crerating an ECERT atribute
        ecert = ET.SubElement(envelope, 'ECERT')
        # Создание дочерних атрибутов элемента ECERT и устновка их значений из xlsx файла
        # Creating ECERT child atributes and setting the values from xlsx file
        for xmlText in xml2excelDict.keys():
            element = ET.SubElement(ecert, xmlText)
            element.text = row[xml2excelDict[xmlText]]

    # Настройка выходного файла и сохранение созданной xml структуры в файл
    # Configuring the output file and saving the created xml structure to a file
    ET.indent(root, '      ')
    tree = ET.ElementTree(root)
    tree.write(outputFileName, encoding="utf-8", xml_declaration=True)


def task2(inputFileName = 'test_input.xlsx', outputFileName='output_xml_v2.xml'):
    '''
    The task2 function opens an xlsx file, generates an xml file of the given
    structure with the information from an xlsx file, and adds a new atribute named SVALUEUSD, which is calculated as the
    ratio of the SVALUE atribute and the exchange rate of the USA Dollar on the SB date.

    Parameters:
    -----------

    inputFileName - the name of the input xlsx file, default values is test_input.xlsx, str

    outputFileName - the name of the output xml file, default values is output_xml_v2.xml, str

    '''
    # Чтение xlsx файла
    # Reading xlsx file
    data = pd.read_excel('test_input.xlsx', header=None, dtype=str)

    # Считывание информации из первых трех строк в отдельные переменные
    # Reading information from the first three lines into separate variables
    proccessingDateText, fileNumberText, fileNameText = data.iloc[:3, 1]

    # Удаление первых четырех строчек таблицы
    # Deleting the first four lines of the table
    data.columns = list(data.iloc[4])
    data = data.iloc[5:, :].reset_index(drop=True)

    # Предобработка данных
    # Data preprocessing

    # Добавление кавычек значениям столбца Client
    # Adding the quotation marks to the values in the Client column
    data['Client'] = data['Client'].map(lambda x: "\"" + x + "\"")

    # Перевод дат в нужный формат
    # Converting dates to the desired format
    data['Issuance Date'] = pd.to_datetime(data['Issuance Date']).dt.strftime('%Y-%m-%d')
    data['SB Date'] = pd.to_datetime(data['SB Date']).dt.strftime('%Y-%m-%d')

    # Перевод IE Code в формат с 10 цифрами, как в Excel файле
    # Converting IE Code to a 10-numbers format, as in an Excel file
    data['IE Code'] = data['IE Code'].str.zfill(10)

    # Словарь соответствий названий атрибутов в xml файле и xlsx файле (если в будущем названия атрибутов поменяются
    # или нужно будет работать с новыми атрибутами, эти изменения можно учесть в словаре)
    # Dictionary of matches of attribute names in an xml file and in an xlsx file (if in the future the attribute names change
    # or it will be neccesary to work with new attributes, these changes can be taken into account in this dictionary)
    xml2excelDict = {'CERTNO': 'Ref no',
                     'CERTDATE': 'Issuance Date',
                     'STATUS': 'Status',
                     'IEC': 'IE Code',
                     'EXPNAME': 'Client',
                     'BILLID': 'Bill Ref no',
                     'SDATE': 'SB Date',
                     'SCC': 'SB Currency',
                     'SVALUE': 'SB Amount'}

    # Создание элементов xml файла
    # Creating elements for an xml file

    # Создание атрибута CERTDATA
    # Creating a CERTDATA atribute
    root = ET.Element('CERTDATA')

    # Создание атрибута FILENAME и установка его значения данными из исходного xlsx файла
    # Creating a FILENAME atribute and setting its values from xlsx file
    filename = ET.SubElement(root, 'FILENAME')
    filename.text = fileNameText

    # Создание атрибута ENVELOPE
    # Crerating an ENVELOPE atribute
    envelope = ET.SubElement(root, 'ENVELOPE')

    # Цикл, в котором данные из xlsx файла добавляются в атрибуты xml файла
    # A loop in which data from an xlsx file is added to the attributes of an xml file
    for iter, row  in data.iterrows():
        # Создание атрибута ECERT
        # Crerating an ECERT atribute
        ecert = ET.SubElement(envelope, 'ECERT')
        # Создание дочерних атрибутов элемента ECERT и устновка их значений из xlsx файла
        # Creating ECERT child atributes and setting the values from xlsx file
        for xmlText in xml2excelDict.keys():
            element = ET.SubElement(ecert, xmlText)
            element.text = row[xml2excelDict[xmlText]]

        # Считывание даты текущего элемента ECERT извлечение значения курса Доллара США в этот день с сайта ЦБ
        # Reading the date of the current ECERT element
        date = dt.strptime(row[xml2excelDict['SDATE']], '%Y-%m-%d').strftime('%d/%m/%Y')
        # Извлечение значения курса Доллара США в этот день с сайта ЦБ РФ
        # И перевод в тип float для дальнейшего выполнения арифметических операций
        # Extracting the value of the US Dollar exchange rate on the current day from the Central Bank's website
        # And converting to the float type for perfoming arithmetic operations
        url = 'http://www.cbr.ru/scripts/XML_daily.asp?date_req=' + date
        currency = pd.read_xml(url, encoding='cp1251')
        USDvalue = float([*currency[currency.ID == 'R01235'].Value][0].replace(',', '.'))

        # Считывание значения атрибута SVALUE текущего элемента ECERT
        # И перевод в тип float для дальнейшего выполнения арифметических операций
        # Reading a SVALUE atribute of the current ECERT element
        # And converting to the float type for perfoming arithmetic operations
        svalue = float(row[xml2excelDict['SVALUE']])

        # Создание атрибута SVALUEUSD и установка его значения частным атрибута SVALUE и курса Доллара США на этот день
        # Результат округляется до двух знаков после запятой
        # Creating an SVALUEUSD atribute and setting its values as the ratio of the SVALUE stribute and the value of the US Dollar exchange rate on the current day
        # The result is rounded to two decimal places
        element = ET.SubElement(ecert, 'SVALUEUSD')
        element.text = f'{svalue / USDvalue:.2f}'

    # Настройка выходного файла и сохранение созданной xml структуры в файл
    # Configuring the output file and saving the created xml structure to a file
    ET.indent(root, '      ')
    tree = ET.ElementTree(root)
    tree.write(outputFileName, encoding="utf-8", xml_declaration=True)
