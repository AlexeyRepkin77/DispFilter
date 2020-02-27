import pandas as pd
from ParseExcelMIS import sliceWorkSheetMIS
from ParseExcelOMS import omsFrame


# Опция отвечающая за вывод всех колонок в консоли
pd.set_option('display.expand_frame_repr', False)

# Указываем место расположения файла
generalArchiveSave = r"E:\Архивы январь февраль\Общий Январь.xlsx"

# Объеденияем два DataFrame (omsFrame и sliceWorkSheetMIS)
genFrame = omsFrame.merge(sliceWorkSheetMIS, how='outer')

# Выбираем столбцы которые нам необходими из общего DataFrame
genFrame = genFrame[['Фамилия', 'Имя', 'Отчество', 'Дата рождения', 'Врач', 'Адрес', 'Тип']]

# Сортируем данные
genFrame = genFrame.sort_values(by=['Фамилия', 'Имя', 'Отчество', 'Дата рождения'])

# Удаляем дубликаты. Параметр keep=False говорит о том что необходимо удалить все дубликаты
genFrame = genFrame.drop_duplicates(subset=['Фамилия', 'Имя', 'Отчество', 'Дата рождения'], keep=False)

# Сохраняем файл
genFrame.to_excel(generalArchiveSave, columns=['Фамилия', 'Имя', 'Отчество', 'Дата рождения', 'Тип', 'Врач', 'Адрес'], index=False)
