import pandas as pd


# Опция отвечающая за вывод всех колонок в консоли
pd.set_option('display.expand_frame_repr', False)

# Указываем место расположения файла
sliceWorkSheetMIS = r"E:\WorkPython\МИС Февраль.xlsx"
pathFileSaveMIS = r"E:\Test\Автоматический вариант МИС.xlsx"

# Читаем данные из файлов с помощью панды указывая что 5 строка это шапка и выбирая колонки C:E, I
sliceWorkSheetMIS = pd.read_excel(sliceWorkSheetMIS, header=5, usecols="C:E, I")

# Выбираем столбец "Пациент" убираем . в инициалах и разбиваем его на отдельные строки
parseFIO_MIS = sliceWorkSheetMIS['Пациент'].str.replace('.', ' ').str.split(' ', expand=True)

# в переменную sliceWorkSheetMIS добавляем полученные столбцы
sliceWorkSheetMIS = [sliceWorkSheetMIS,parseFIO_MIS]

# Создаём DataFrame из листа sliceWorkSheetMIS и убираем пустые строки
sliceWorkSheetMIS = pd.concat(sliceWorkSheetMIS, axis=1)

# Создаём список для фильтра для поля(колонке) Врач, закрывший карту МО
keysFilter = ['016 Кибалова Э.Т.',
              '020 Никишина Е.Э.',
              '19 Джураева С.Т.',
              '15 Ваничкина И.Н.',
              '42 Глушаева О.Ф.',
              '06 Романенко Г.Б.']

# Даём новые имена всем полученным столбцам в фрейме
sliceWorkSheetMIS.columns = ['ФИО', 'Дата рождения', 'Адрес', 'Врач', 'Фамилия', 'Имя', 'Отчество', 'Удалить']

# Добавляем постоянный столбец и заполняем данными "МИС"
sliceWorkSheetMIS['Тип'] = 'МИС'

# Применяем фильтр для колонки "Врач"
sliceWorkSheetMIS = sliceWorkSheetMIS.loc[sliceWorkSheetMIS['Врач'].isin(keysFilter)]

# Количество строк в файле
countLineMIS = len(sliceWorkSheetMIS.index)

# Сохраняем в Excel необходимые колонки отрезая индексы и шапку
#sliceWorkSheetMIS.to_excel(pathFileSaveMIS, columns=['Фамилия', 'Имя', 'Отчество', 'Дата рождения', 'Врач', 'Адрес', 'Тип'], index=False, header=False)
print(countLineMIS)