import pandas as pd


# Опция отвечающая за вывод всех колонок в консоли
pd.set_option('display.expand_frame_repr', False)



def my_Func_Parse_OMS(LocationExcelFileOMS):
    sliceWorkSheetOMS = pd.read_excel(LocationExcelFileOMS, header=5, usecols = "A:D, I, J")



    # Разбиваем на отдельные колонки колонку ФИО
    parseFIO_OMS = sliceWorkSheetOMS['ФИО пациента'].str.split(' ', expand=True)

    #Оставляем первую букву от Имени
    parseFIO_OMS[1] = parseFIO_OMS[1].str[0]
    parseFIO_OMS[2] = parseFIO_OMS[2].str[0]

    # Склеиваем двае переменные методом concat. Параметр axis=1 говорит о том что необходимо добавить с первйо строки а не в конце файла
    omsFrame = pd.concat([sliceWorkSheetOMS, parseFIO_OMS], axis=1)

    # Создаём список для фильтра для поля(колонке) Специальность / Отделение
    keysFilter = ['ДИСП. 76 лет и > женщины',
                'ДИСП. 55 лет мужчины',
                'ДИСП. 41-63 года женщины',
                'ДИСП. 40-64 года женщины',
                'ДИСП. 65-75 лет женщины',
                'ДИСП. 18-39 лет женщины',
                'ДИСП. 66-74 лет женщины',
                'ДИСП. 40-62 года мужчины',
                'ДИСП. 76 лет и > мужчины',
                'ДИСП. 41-63 года мужчины',
                'ДИСП. 50,60,64 года мужчины',
                'ДИСП. 65-75 лет мужчины',
                'ДИСП. 18-39 лет мужчины',
                'ДИСП. 45 лет женщины',
                'ДИСП. 45 лет мужчины']

    # Фильтруем строки по заданным ключам
    omsFrame = omsFrame.loc[omsFrame['Специальность / Отделение'].isin(keysFilter)]

    # Добавляем постоянный столбец и заполняем данными "ОМС"
    omsFrame['Тип'] = 'ОМС'

    # Убираем значок № из файла
    for column in omsFrame['Полис']:
        omsFrame['Полис'] = omsFrame['Полис'].str.replace('№', '')

    # Выбираем из фрейма необходимые столбцы
    omsFrame = omsFrame[[0, 1, 2, 'Дата рождения', 'Врач', 'Тип']]

    # Переменовываем их
    omsFrame.columns = ['Фамилия', 'Имя', 'Отчество','Дата рождения', 'Врач', 'Тип']

    # Количество строк в файле
    countLineOMS = len(omsFrame.index)


    print(countLineOMS)


# #omsFrame.to_excel(pathFileSaveOMS, index=False, header=False,)