import openpyxl
from PyQt5 import QtWidgets
import sys  # нужен для передачи argv в QApplication
import os   # для работы с файлами
import pandas   # для файлов exel
import one  # Это наш конвертированный файл дизайна 1
import two  # Это наш конвертированный файл дизайна 2
from datetime import date
import docx
from dateutil.parser import parse


def writeinfile(file_data, file_name):
    check_file = os.path.exists(file_name)
    if check_file:
        with open(file_name, 'a') as filer:
            filer.writelines(file_data + '\n')
    else:
        with open(file_name, 'w') as filer:
            filer.write(file_data + '\n')


def error():
    reply = QtWidgets.QMessageBox()
    reply.setWindowTitle("Ошибка")
    reply.setIcon(reply.Warning)
    reply.setText("ОШИБКА! \n Неверный формат!")
    reply.addButton("Ok", QtWidgets.QMessageBox.ButtonRole.RejectRole)
    reply.exec_()


class MyWindow(QtWidgets.QDialog, two.Ui_Form):
    def __init__(self):
        self.book = openpyxl.Workbook()  # Создаём файл экселя
        self.last_row = 0
        super(MyWindow, self).__init__()
        self.setupUi(self)  # Это нужно для инициализации дизайна
        self.to_print_list.clicked.connect(self.to_print)
        self.add_to_report.clicked.connect(self.add_report)
        self.create_report.clicked.connect(self.create_rep)
        self.close_pr.clicked.connect(self.close_form)  # Закрытие окна по нажатию кнопки
        self.save_list.clicked.connect(self.save_in_file)

    def save_in_file(self):
        reply = QtWidgets.QMessageBox()
        reply.setWindowTitle("Сохранение файла")
        reply.setText("Выберете формат сохранения файла")
        reply.addButton("docx", QtWidgets.QMessageBox.AcceptRole)
        reply.addButton("xlsx", QtWidgets.QMessageBox.AcceptRole)
        reply.addButton("Отмена", QtWidgets.QMessageBox.ButtonRole.RejectRole)
        result = reply.exec_()
        myfile = ''
        way = ''
        if result == 1:
            myfile = openpyxl.Workbook()
            sht = myfile.active
            for i in range(len(self.list_absent) - 1):
                sht['A' + str(i + 1)] = self.list_absent.item(i).text()
            way = QtWidgets.QFileDialog.getSaveFileName(self, "Save file", '*.xlsx', '*.xlsx')[0]
        elif result == 0:
            myfile = docx.Document()
            for i in range(len(self.list_absent) - 1):
                myfile.add_paragraph(self.list_absent.item(i).text())
            way = QtWidgets.QFileDialog.getSaveFileName(self, "Save file", '*.docx', '*.docx')[0]
        if way:
            myfile.save(way)  # Сохраняем файл

    # Корректное закрытие программы
    def close_form(self):
        try:
            os.remove("fileRemember.txt")  # Удаляем файл, т.к. на этом этапе корректное завершение программмы
            os.remove("to_print.txt")   # Удаляем файл для печати недостающих дел
        except FileNotFoundError:
            pass
        self.close()

    def to_print(self):
        for i in range(len(self.list_absent)-1):
            writeinfile(self.list_absent.item(i).text(), 'to_print.txt')
        os.startfile('to_print.txt', 'print')

    # Функция добавления дел в отчёт
    def add_report(self):
        sheet = self.book.active
        for i in range(len(self.list_absent)-1):
            sheet['B' + str(self.last_row)] = self.list_absent.item(i).text()
            sheet['A' + str(self.last_row)] = self.last_row - 1
            self.last_row += 1
        reply = QtWidgets.QMessageBox()
        reply.setWindowTitle("Создание отчёта")
        reply.setText("Дела успешно добавлены в отчёт")
        reply.setStandardButtons(QtWidgets.QMessageBox.StandardButton.Ok)
        reply.exec_()

    def create_rep(self):
        # Вызываем диалог для сохранения отчёта
        way = QtWidgets.QFileDialog.getSaveFileName(self, "Save file", '/МФЦ ' + str(date.today()), '*.xlsx')[0]
        if way:
            self.book.save(way)  # Сохраняем файл
            self.book.close()    # закрываем файл

    # Функция формирования отчёта
    def report(self, mas):
        sheet = self.book.active
        sheet.title = 'Лист1'  # Задаём название листа
        # прописываем заголовки
        sheet['A1'] = 'п / п'
        sheet['B1'] = '№ Заявления'
        sheet['C1'] = '№ Акта / Дата завершения'
        count = 2  # Задаём строку, с которой заполнится файл. 2 строка
        '''for i in range(len(mas)):
            if i % 2 == 0:
                num = mas[i]
            else:
                # Заполняем файл
                for k in range(len(mas[i])):
                    sheet['B' + str(count)] = mas[i][k]
                    sheet['C' + str(count)] = str(num) + " / " + str(date.today())
                    sheet['A' + str(count)] = count - 1
                    count += 1
        sheet.column_dimensions['B'].width = 30
        sheet.column_dimensions['C'].width = 20
        self.last_row = count   # Передаём число последней строки, чтобы записать недостающие дела в конец списка
        '''
        for i in range(len(mas)):
            # Заполняем файл
            for k in range(len(mas[i][1])):
                sheet['B' + str(count)] = mas[i][1][k]
                sheet['C' + str(count)] = str(mas[i][0]) + " / " + str(mas[i][2][k])
                sheet['A' + str(count)] = count - 1
                count += 1
                # print(count)
        sheet.column_dimensions['B'].width = 30
        sheet.column_dimensions['C'].width = 20
        self.last_row = count  # Передаём число последней строки, чтобы записать недостающие дела в конец списка


class ExampleApp(QtWidgets.QMainWindow, one.Ui_MainWindow):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        self.array_akt = []   # массив для хранения списка всех актов
        self.array_pakages = []     # массив для пакетов
        self.mas3 = []   # Инициализируем список для отчёта 3х мерный массив
        self.number_otdel = []   # код отдела
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации дизайна
        self.packege_file.clicked.connect(self.browse_folder)  # выбрать файл Пакетов для сверки
        self.akt_files.clicked.connect(self.browse_folder_akt)  # загрузить файлы с Актами
        self.compare.clicked.connect(self.compare_akt)  # обработка кнопки сравнения
        self.display.clicked.connect(self.compare_display)  # обработка кнопки "отобразить"
        self.mywin = MyWindow()     # экземпляр класса втрого интерфейса
        self.readfromfile()
        self.select_otdel()
        # self.changeFont()
        self.mywin.close_pr.clicked.connect(self.close)  # Закрытие окна по нажатию кнопки
        self.tab2_akt.itemDoubleClicked.connect(self.remove_item)  # при двойном клике удаляется путь к файлу-акту
        self.mywin.list_absent.itemDoubleClicked.connect(self.find_people)

    def find_people(self):
        pep = self.mywin.list_absent.selectedItems()
        for item in pep:
            number_find = self.mywin.list_absent.currentItem().text()
            print(number_find)
            directory = self.tab1_way.text()  # берём путь к файлу из окна интерфейса 1
            directory = directory.replace("/", "\\")  # Замена символов, т.к. путь к файлу питон считывает по-своему
            # поиск по параметрам в таблице
            excel_data = pandas.read_excel(directory, sheet_name='Лист1',
                                           usecols=['Номер внешней системы', 'Провел Удостоверение'])
            mas = []  # Для удобства создаём массив
            # Выполняем сортировку файла
            sort = excel_data[(excel_data['Номер внешней системы'] == number_find)]
            mas = sort['Провел Удостоверение'].tolist()
            print(mas[0])
            answer = QtWidgets.QMessageBox()
            answer.setWindowTitle("Провел Удостоверение")
            answer.setIcon(answer.Icon.Information)
            answer.setText(f"Провел Удостоверение: {str(mas[0])}")
            answer.addButton("Ok", QtWidgets.QMessageBox.ButtonRole.RejectRole)
            answer.exec_()


    # Функция для удаление пути к файлу-акту
    def remove_item(self):
        stitems = self.tab2_akt.selectedItems()
        if stitems:
            for item in stitems:
                self.tab2_akt.takeItem(self.tab2_akt.row(item))
        self.tab2_list.clear()
        # self.report = []
        self.array_akt = []

    '''def font(self) -> QtWidgets.QMainWindow.QFont:
        QtWidgets.QMessageBox.'''

    '''def changeFont(self):
        if self.menu_3.actions()[0].isChecked():
            print("FONT")
            # для замены шрифта
            font, ok = QtWidgets.QFontDialog.getFont()
            if ok:
                print(font)'''

    def readfromfile(self):
        try:
            with open('fileRemember.txt', 'r') as filer:
                arr = filer.read().splitlines()
                for i in range(len(arr)):
                    if "ПАКЕТЫ" in arr[i]:
                        self.tab1_way.setText(arr[i])
                    else:
                        self.tab2_akt.addItem(arr[i])
        except FileNotFoundError:
            # print("no file")
            pass

    def select_otdel(self):     # Функция выбора отдела
        # Задаём код каждому отделу
        self.shelcovo.setData(50.158)   # Щелковский отдел
        self.klinski.setData(50.124)    # Клинский отдел
        self.balashiha.setData([50.110, 50.420])    # Отдел по г. Балашиха
        self.ogku_grp.setData(50.215)   # Отдел государственного кадастрового учета и государственной регистрации прав
        self.VolkLotoshShah.setData([50.111, 50.130, 50.157])   # Волоколамскому, Лотошинскому и Шаховскому районам
        self.VoskrKolomLuh.setData([50.112, 50.125, 50.417])    # Воскресенскому, Коломенскому и Луховицкому районам
        self.DubnaTald.setData([50.116, 50.153])    # Отдел по г. Дубна и Талдомскому району
        self.JuckRamensk.setData([50.119, 50.145])  # Отдел по г. Жуковский и Раменскому району
        self.leninski.setData(50.128)   # Ленинский отдел
        self.LobnDmitr.setData([50.113, 50.129])    # Отдел по г. Лобня и Дмитровскому району
        self.ElectrstNoginsk.setData([50.137, 50.706])  # Отдел по г. Электросталь и Ногинскому району
        self.pushkinski.setData(50.144)  # Пушкинский отдел
        self.EgorShatur.setData([50.117, 50.156])  # Отдел по Егорьевскому и Шатурскому районам
        # Отдел по Зарайскому, Каширскому, Озерскому и Серебряно-Прудскому районам
        self.ZaraiKashOzerSerg.setData([50.120, 50.123, 50.140, 50.149])
        self.MojRuzovski.setData([50.134, 50.147])  # Отдел по Можайскому и Рузскому районам
        self.Narofominsk.setData(50.136)  # Наро-Фоминский отдел
        self.OrZuPavlPas.setData(50.141)  # Отдел по Орехово-Зуевскому и Павлово-Посадскому районам
        self.PodChehovsk.setData([50.143, 50.155])  # Отдел по Подольскому и Чеховскому районам
        self.SerpStupinski.setData([50.150, 50.152])  # Отдел по Серпуховскому и Ступинскому районам
        self.opguElectr.setData(50.215)  # Отдел предоставления государственных услуг в электронном виде
        self.sergposad.setData(50.148)  # Сергиево - Посадский отдел
        self.otdel1.setData([50.133, 50.414, 50.415, 50.419])  # Территориальный отдел № 1
        self.otdel2.setData([50.115, 50.421])  # Территориальный отдел № 2
        self.otdel_50_422.setData(50.422)  # Территориальный отдел (новый ЕСТО 50.422)
        self.sun.setData(50.416)  # Солнечногорский отдел

        find = self.menu_4.actions()    # Получаем список всех отделов из меню
        # создаём группу для списка отделов
        action_group = QtWidgets.QActionGroup(self)
        for i in range(len(find)):
            action_group.addAction(find[i])
        action_group.setExclusive(True)     # чтобы выбрать можно было только один отдел

        # если зажата "ввести вручную", то вызываем окно для ввода кода отдела
        if self.handswrite.isChecked():
            text, ok = QtWidgets.QInputDialog.getText(self, 'Выбор отдела', 'Введите код отдела:')
            if ok:
                try:
                    text = float(text)
                    self.number_otdel = text
                except ValueError:
                    error()
                    self.select_otdel()
        else:   # если не зажата "ввести вручную", проверяем другие пункты
            for i in range(len(find)):
                if find[i].isChecked():  # Если зажата галочка на отделе выбираем этот отдел
                    self.number_otdel = find[i].data()  # Получаем код отдела из списка определённого выше

    # функция бработки ПАКЕТов ДЛЯ СВЕРКИ
    def browse_folder(self):
        self.tab1_pakages.clear()  # На случай, если в списке уже есть элементы, чистим таблицы
        self.tab1_way.clear()
        try:
            directory = QtWidgets.QFileDialog.getOpenFileName(self, 'Open file', '*.xlsx')[0]
            # открыть диалог выбора директории и установить значение переменной равной пути к выбранной директории
            if directory:  # не продолжать выполнение, если не выбрана директория
                self.tab1_way.setText(directory)
                writeinfile(directory, 'fileRemember.txt')  # Записываем путь в файл
                self.compare_display()  # Вызываем функцию отображения на экран
                '''====================================

                # поиск по параметрам в таблице
                excel_data = pandas.read_excel(directory, sheet_name='Лист1',
                                               usecols=['Номер внешней системы', 'Провел ППЭ ЕСТО'])
                sort = 0
                mas = []
                # Выполняем сортировку файла
                if type(number_otdel) is list:  # Если у отдела несколько кодов
                    for i in range(len(number_otdel)):
                        sort = excel_data[(excel_data["Провел ППЭ ЕСТО"] == number_otdel[i])]
                        mas.extend(sort['Номер внешней системы'].tolist())
                else:   # Иначе если код это одно число
                    sort = excel_data[(excel_data["Провел ППЭ ЕСТО"] == number_otdel)]
                    mas = sort['Номер внешней системы'].tolist()
                # Выводим на экран данные
                self.tab1_pakages.addItem('Общее количество: ' + str(len(mas)))
                for i in range(len(mas)):
                    self.tab1_pakages.addItem(mas[i]) # выводим на экран дела
                self.array_pakages = mas

                =========================================='''
        except ValueError:
            error()
            # self.tab1_pakages.addItem("!!!Ошибка!!! \nНеверный формат! \nЗагрузите файл снова!")
            """
            error = QtWidgets.QMessageBox()
            #error.setWindowTitle("Ошибка")
            error.setText('Неверный формат')
            #error.setIcon(QtWidgets.QMessageBox.Warning)
            error.setStandardButtons(QtWidgets.QMessageBox.ok|QtWidgets.QMessageBox.Cancel)
            error.exec_()
            """

    # функция обработки файлов с АКТами
    def browse_folder_akt(self):
        # self.tab2_list.clear()  # На случай, если в списке уже есть элементы, чистим таблицу
        try:
            # вызов диалога для выбора файлов
            file = QtWidgets.QFileDialog.getOpenFileNames(self, 'Open file', '*.xlsx')[0]
            if file:  # не продолжать выполнение, если не выбрана директория
                for i in range(len(file)):
                    self.tab2_akt.addItem(file[i])  # отображаем путь файла на экране
                    writeinfile(file[i], 'fileRemember.txt')  # Записываем путь в файл
                self.compare_display()
        except ValueError:
            error()
            self.tab2_akt.takeItem(len(self.tab2_akt)-1)    # Удаляем путь, если файл не соответствует формату

    # вывод на экран. обработка файлов
    def compare_display(self):
        if self.tab1_pakages.count() == 0:
            self.tab1_pakages.clear()
            # Если в списке "пакеты для сверки" ещё не отображены дела, выполняем код
            # if self.tab1_pakages.count() == 0:
            directory = self.tab1_way.text()    # берём путь к файлу из окна интерфейса 1
            directory = directory.replace("/", "\\")    # Замена символов, т.к. путь к файлу питон считывает по-своему
            if directory:
                self.select_otdel()  # вызываем функцию для определения кода отдела
                # поиск по параметрам в таблице
                excel_data = pandas.read_excel(directory, sheet_name='Лист1',
                                               usecols=['Номер внешней системы', 'Провел ППЭ ЕСТО'])
                mas = []    # Для удобства создаём массив
                # Выполняем сортировку файла
                if type(self.number_otdel) is list:  # Если у отдела несколько кодов
                    for i in range(len(self.number_otdel)):
                        sort = excel_data[(excel_data["Провел ППЭ ЕСТО"] == self.number_otdel[i])]
                        mas.extend(sort['Номер внешней системы'].tolist())
                else:  # Иначе если код это одно число
                    sort = excel_data[(excel_data["Провел ППЭ ЕСТО"] == self.number_otdel)]
                    mas = sort['Номер внешней системы'].tolist()
                # Выводим на экран данные
                self.tab1_pakages.addItem('Общее количество: ' + str(len(mas)))
                for i in range(len(mas)):
                    self.tab1_pakages.addItem(mas[i])  # выводим на экран дела
                self.array_pakages = mas    # Передаём список "пакетов для сверки" в локальную переменную класса

        self.mas3 = [[] * 3 for _ in range(len(self.tab2_akt))]
        self.array_akt = []
        for i in range(len(self.tab2_akt)):
            string = self.tab2_akt.item(i).text()
            number_akt = string[string.rindex("-")+1:-5]    # Получем № акта, для дальнейшего добавления в отчёт
            self.mas3[i].append(number_akt)
            # self.report.append(int(number_akt))     # добавляем номер акта в массив для отчёта
            wer = []       # делаем список чтобы добавлять в него дела из текущего файла

            # считываем данные из файлов по параметрам
            data_file = pandas.read_excel(self.tab2_akt.item(i).text(), sheet_name=0, usecols="B:C", header=None,
                                          dtype=str)
            arr = data_file[1].tolist()     # дела. добавляем столбик В в список
            arrv = data_file[2].tolist()    # дата. добавляем столбик C в список
            for k in range(len(arr)):
                if 'MFC' in str(arr[k]) and arr[k][-1].isdigit():
                    if arr[k] in self.array_akt:
                        continue  # Если файл в списке уже есть - пропустить его
                    self.array_akt.append(arr[k])
                    self.tab2_list.addItem(arr[k])  # выводим список на экран
                    wer.append(arr[k])
            self.mas3[i].append(wer)    # добавляем список wer в массив с отчётом
            begin = 0
            end = 0
            for k in range(len(arrv)):
                if '№ КУВД' in str(arrv[k]):
                    begin = k + 1
                if k > begin != 0 and str(arrv[k]) == 'nan':
                    end = k
                    break
            dataakt = arrv[begin:end]
            for dt in range(len(dataakt)):
                if len(dataakt[dt]) > 10:
                    dataakt[dt] = dataakt[dt][:10]
                    dataakt[dt] = dataakt[dt][:10]
                    d = parse(dataakt[dt])
                    dataakt[dt] = d.strftime('%d.%m.%Y')
            self.mas3[i].append(dataakt)

    # функция сравнения файлов
    def compare_akt(self):
        # self.list_absent.addItem(self.array_akt)
        self.mywin.list_absent.clear()
        self.mywin.show()   # вызываем 2 интерфейс
        for i in range(len(self.array_pakages)):
            if self.array_pakages[i] not in self.array_akt:
                # если файлы не совпадают в списках
                self.mywin.list_absent.addItem(self.array_pakages[i])   # Выводим список недостающих дел на экран
        self.mywin.list_absent.addItem('Общее количество: ' + str(len(self.mywin.list_absent)))
        self.mywin.report(self.mas3)


def main():
    appy = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    appy.exec_()  # и запускаем приложение


if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()

    '''
    чтобы сделать EXE-файл
    Пишем команду pyinstaller -F -w -i( to set up icon on your .exe) main.py , где main.py — ваш python скрипт. 
    Вот что означает каждый флаг: F – этот флаг отвечает за то, чтобы в созданной папке dist, в которой и будет 
    храниться наш исполняемый файл не было очень много лишних файлов, модулей и т.п. -w – этот флаг вам понадобится 
    в том случае, если приложение использует GUI библиотеки ( tkinter , PyQt5 , т.п.), оно блокирует создание 
    консольного окна, если же ваше приложение консольное, вам этот флаг использовать не нужно. -i – этот флаг отвечает 
    за установку иконки на наш исполняемый файл, после флага нужно указать полный путь к иконке с указанием её имени. 
    Например: D:\\LayOut\\icon.ico
    '''