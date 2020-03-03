import pandas as pd
import ParseExcelMIS
from ParseExcelOMS import my_Func_Parse_OMS
from ParseExcelMIS import my_Func_Parse_MIS

# Опция отвечающая за вывод всех колонок в консоли
pd.set_option('display.expand_frame_repr', False)

def my_Func_Dublicate(omsFrame, LocationExcelFileMIS):
    omsFrame = pd.read_excel(omsFrame, header=5, usecols="A:D, I, J")
    sliceWorkSheetMIS = pd.read_excel(LocationExcelFileMIS, header=5, usecols="C:E, I")
    genFrame = omsFrame.merge(sliceWorkSheetMIS, how='outer')

    # Объеденияем два DataFrame (omsFrame и sliceWorkSheetMIS)
    genFrame = my_Func_Parse_OMS.omsFrame.merge(my_Func_Parse_MIS.sliceWorkSheetMIS, how='outer')

    # Выбираем столбцы которые нам необходими из общего DataFrame
    genFrame = genFrame[['Фамилия', 'Имя', 'Отчество', 'Дата рождения', 'Врач', 'Адрес', 'Тип']]

    # Сортируем данные
    genFrame = genFrame.sort_values(by=['Фамилия', 'Имя', 'Отчество', 'Дата рождения'])

    # Удаляем дубликаты. Параметр keep=False говорит о том что необходимо удалить все дубликаты
    genFrame = genFrame.drop_duplicates(subset=['Фамилия', 'Имя', 'Отчество', 'Дата рождения'], keep=False)

    # Сохраняем файл
    #genFrame.to_excel(generalArchiveSave, columns=['Фамилия', 'Имя', 'Отчество', 'Дата рождения', 'Тип', 'Врач', 'Адрес'], index=False)
    print(genFrame)
