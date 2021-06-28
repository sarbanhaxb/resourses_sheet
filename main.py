import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox
import docx
from UI import Ui_main_window
import pandas as pd
import openpyxl
from docx.shared import Inches, Pt

class resourses_sheet(QMainWindow):
    def __init__(self):
        super(resourses_sheet, self).__init__()
        self.UI = Ui_main_window()
        self.UI.setupUi(self)
        self.init_UI()

    def make_magic(self, fname):
        book = pd.read_excel(fname, engine='openpyxl')
        column_one = book.columns.tolist()[0]
        column_two = book.columns.tolist()[1]
        column_three = book.columns.tolist()[2]
        column_four = book.columns.tolist()[3]
        num_slice = []
        for i in range(book.shape[0]):
            if str(book[column_one][i]).replace(" ", '').lower() == 'материалы' or str(book[column_two][i]).replace(" ", '').lower() == 'материалы' or str(
                    book[column_three][i]).replace(" ", '').lower() == 'материалы':
                num_slice.append(i)
            elif str(book[column_one][i]).replace(" ", '').lower() == 'итого"материалы"' or str(
                    book[column_two][i]).replace(" ", '').lower() == 'итого"материалы"':
                num_slice.append(i)
                break
        ws = []
        for i in range(num_slice[0] + 1, num_slice[1]):
            ws.append([str(book[column_one][i]), str(book[column_two][i]), str(book[column_three][i]),
                       str(book[column_four][i])])
        df = pd.DataFrame(ws, columns=['code', 'name', 'unit', 'amount'])

        electrodes = 0
        paints = 0
        geosynthetic_nt = 0
        geogrids_nt = 0
        thermal = 0
        bituminous = 0
        sands = 0
        geomembranes = 0
        slabs = 0
        semens = 0

        for index, row in df.iterrows():
            """расчет электродов"""
            if str(row['name']).lower().find('электрод') == 0:
                if str(row['unit']).lower() == 'кг':
                    electrodes += float(row['amount']) / 1000
                elif str(row['unit']).lower() == 'т':
                    electrodes += float(row['amount'])
            """расчет геополотна нетканого"""
            if (str(row['code']).lower().find('фссц-01.7.12.05') == 0 or str(row['code']).lower().find('01.7.12.05') == 0) and (str(row['name']).lower().find('нетканый') == 0 or str(row['name']).lower().find('геополотно') == 0 or str(row['name']).lower().find('геотекстиль') == 0 or str(row['name']).lower().find('дренажный композит') == 0 or str(row['name']).lower().find('материал геотекстильный нетканый') == 0 or str(row['name']).lower().find('полотно иглопробивное') == 0 or str(row['name']).lower().find('полотно нетканное ') == 0):
                if str(row['unit']).lower() == 'м2':
                    geosynthetic_nt += float(row['amount'])
            """расчет георешетки нетканой """
            if (str(row['code']).lower().find('фссц-01.7.12.07') == 0 or str(row['code']).lower().find('01.7.12.07') == 0) and str(row['name']).lower().find('георешетка ') == 0:
                if str(row['unit']).lower() == 'м2':
                    geogrids_nt += float(row['amount'])
            """расчет геомембраны 01.7.12.04"""
            if (str(row['code']).lower().find('01.7.12.04') == 0 or str(row['code']).lower().find('фссц-01.7.12.04') == 0) and (str(row['name']).lower().find('геомембрана') == 0):
                if str(row['unit']).lower() == 'м3':
                    geomembranes += float(row['amount'])
            """расчет теплоизоляционного материала"""
            if (str(row['code']).lower().find('фссц-12.2.04') == 0 or str(row['code']).lower().find('12.2.04') == 0) and (str(row['name']).lower().find('маты') == 0 or str(row['name']).lower().find('пакеты') == 0):
                if str(row['unit']).lower() == 'м3':
                    thermal += float(row['amount'])
            """расчет битумной мастики"""
            if (str(row['code']).lower().find('01.2.03.03') == 0 or str(row['code']).lower().find('фссц-01.2.03.03') == 0) and (str(row['name']).lower().find('мастика ') == 0 or str(row['name']).lower().find('состав мастичный') == 0):
                if str(row['unit']).lower() == 'кг':
                    bituminous += float(row['amount'])
                elif str(row['unit']).lower() == 'т':
                    bituminous += float(row['amount']) * 1000
            """расчет ЛКМ"""
            #расчет грунтовки группа 14.4.01
            if (str(row['code']).lower().find('14.4.01.') == 0 or str(row['code']).lower().find('фссц-14.4.01.') == 0) and (str(row['name']).lower().find('грунтовка') == 0 or str(row['name']).lower().find('праймер') == 0 or str(row['name']).lower().find('cостав') == 0 or str(row['name']).lower().find('грунт') == 0 or str(row['name']).lower().find('покрытие') == 0):
                if str(row['unit']).lower() == 'кг':
                    paints += float(row['amount'])
                elif str(row['unit']).lower() == 'т':
                    paints += float(row['amount']) * 1000
            #расчет ЛКМ группы 14.4.02
            if (str(row['code']).lower().find('14.4.02.') == 0 or str(row['code']).lower().find('фссц-14.4.02.') == 0) and (str(row['name']).lower().find('белила') == 0 or str(row['name']).lower().find('краска') == 0 or str(row['name']).lower().find('покрытие') == 0 or str(row['name']).lower().find('эмаль') == 0 or str(row['name']).lower().find('композиция') == 0 or str(row['name']).lower().find('топкоут') == 0):
                if str(row['unit']).lower() == 'кг':
                    paints += float(row['amount'])
                elif str(row['unit']).lower() == 'т':
                    paints += float(row['amount']) * 1000
            #расчет лаки группы 14.4.03
            if (str(row['code']).lower().find('14.4.03.') == 0 or str(row['code']).lower().find('фссц-14.4.03.') == 0) and (str(row['name']).lower().find('лак') == 0 or str(row['name']).lower().find('раствор хлорсульфированного') == 0 or str(row['name']).lower().find('нитролак') == 0 or str(row['name']).lower().find('покрытие полиуретановое КТ ') == 0):
                if str(row['unit']).lower() == 'кг':
                    paints += float(row['amount'])
                elif str(row['unit']).lower() == 'т':
                    paints += float(row['amount']) * 1000
            #расчет эмали группы 14.4.04
            if (str(row['code']).lower().find('14.4.04.') == 0 or str(row['code']).lower().find('фссц-14.4.04.') == 0) and (str(row['name']).lower().find('нитроэмаль') == 0 or str(row['name']).lower().find('эмаль') == 0 or str(row['name']).lower().find('покрытие') == 0 or str(row['name']).lower().find('состав (эмаль)') == 0 or str(row['name']).lower().find('шликер') == 0 or str(row['name']).lower().find('спрей-эмаль') == 0 or str(row['name']).lower().find('эпималь') == 0):
                if str(row['unit']).lower() == 'кг':
                    paints += float(row['amount'])
                elif str(row['unit']).lower() == 'т':
                    paints += float(row['amount']) * 1000
            #другое Прайс-лист 2.1
            if str(row['code']).lower().find('прайс-лист') == 0 and (str(row['name']).lower().find('эмаль') != -1):
                if str(row['unit']).lower() == 'кг':
                    paints += float(row['amount'])
                elif str(row['unit']).lower() == 'т':
                    paints += float(row['amount']) * 1000
            """расчет песка"""
            if (str(row['code']).lower().find('02.3.01') == 0 or str(row['code']).lower().find('фссц-02.3.01') == 0 or str(row['code']).lower().find('данные заказчика') == 0) and (str(row['name']).lower().find('песок') == 0):
                if str(row['unit']).lower() == 'м3':
                    sands += float(row['amount'])*1.6
                elif str(row['unit']).lower() == 'кг':
                    sands += float(row['amount'])
                elif str(row['unit']).lower() == 'т':
                    sands += float(row['amount'])*1000
            if str(row['code']).lower().find('цена заказчика') == 0 and (str(row['name']).lower().find('песок') != -1):
                if str(row['unit']).lower() == 'м3':
                    sands += float(row['amount']) * 1.6
                elif str(row['unit']).lower() == 'кг':
                    sands += float(row['amount'])
                elif str(row['unit']).lower() == 'т':
                    sands += float(row['amount'])*1000
            """расчет щебня"""
            """расчет плит дорожных"""
            if str(row['code']).lower().find('05.1.08.06-0063') != -1 and str(row['name']).lower().find('плиты дорожные') == 0:
                if str(row['unit']).lower() == 'шт':
                    slabs += float(row['amount'])*4.2

            """расчет металлоконструкций"""
            """расчет стальных труб"""
            """расчет цементного раствора"""
            """расчет песчано-гравийного раствора"""
            """расчет семян трав"""
            if str(row['code']).lower().find('16.2.02.07') != -1 and str(row['name']).lower().find('семена') == 0:
                if str(row['unit']).lower() == 'кг':
                    semens += float(row['amount'])*4.2
            """расчет удобрений"""
            if str(row['code']).lower().find('16.2.02.07') != -1 and str(row['name']).lower().find('семена') == 0:
                if str(row['unit']).lower() == 'кг':
                    semens += float(row['amount'])*4.2
            """"""

        global materials
        materials = [['Наименование основных строительных конструкций, изделий и материалов', 'Единица измерения', 'Всего']]
        electrodes = ['Электроды', 'кг', round(electrodes*1000, 2)]
        paints = ['Лакокрасочные материалы', 'кг', round(paints, 2)]
        geosynthetic_nt = ['Нетканое геополотно', 'м2', round(geosynthetic_nt, 2)]
        geogrids_nt = ['Нетканая георешетка', 'м2', round(geogrids_nt, 2)]
        thermal = ['Теплоизоляционный материал', 'м3', round(thermal, 2)]
        bituminous = ['Битумно-резиновая мастика', 'кг', round(bituminous, 2)]
        sands = ['Песок', 'кг', round(sands, 2)]
        geomembranes = ['Геомембрана', 'м2', round(geomembranes, 2)]
        slabs = ['Плиты дорожные', 'т', round(slabs, 2)]
        semens = ['Семена трав', "кг", round(semens, 2)]

        if electrodes[2] != 0:
            materials.append(electrodes)
        if paints[2] != 0:
            materials.append(paints)
        if geosynthetic_nt[2] != 0:
            materials.append(geosynthetic_nt)
        if geogrids_nt[2] != 0:
            materials.append(geogrids_nt)
        if thermal[2] != 0:
            materials.append(thermal)
        if bituminous[2] != 0:
            materials.append(bituminous)
        if sands[2] != 0:
            materials.append(sands)
        if geomembranes[2] != 0:
            materials.append(geomembranes)
        if slabs[2] != 0:
            materials.append(slabs)
        if semens[2] != 0:
            materials.append(semens)

    def showDialog(self):
        try:
            fname = QFileDialog.getOpenFileName(self, 'Open file', '', "Все файлы Excel (*.xlsx)")[0]
            try:
                self.make_magic(fname)
                self.UI.label_5.setText('Ведомость загружена')
            except IndexError:
                QMessageBox.critical(self, "Ошибка ", "Выбран неверный формат файла", QMessageBox.Ok)
                self.UI.label_5.setText('')
        except FileNotFoundError:
            pass

    def Save_file(self):
        name_pd = self.UI.name_pd.toPlainText()
        chifr = self.UI.chifr.toPlainText()
        fn = 'Задание отделу ЭиПБ' #название сохраняемого файла
        try:
            doc = docx.Document()
            style = doc.styles['Normal']
            font = style.font
            font.name = 'Times New Romans'
            font.size = Pt(12)
            a = doc.add_paragraph()
            a.add_run('ЗАДАНИЕ ОТДЕЛУ ЭиПБ').bold = True
            a.alignment = 1
            b = doc.add_paragraph(f'Заказ № {chifr}\nСтадия ПД')
            b.alignment = 2
            doc.add_paragraph('')
            table = doc.add_table(rows=3, cols=2, style='Table Grid')
            table.cell(0, 0).text = 'От отдела\n'
            table.cell(0, 1).text = 'СМ'
            table.cell(1, 0).text = 'Отделу\n'
            table.cell(1, 1).text = 'ЭиПБ'
            table.cell(2, 0).text = 'Наименование объекта\n'
            table.cell(2, 1).text = f'{name_pd}'
            table = doc.add_table(rows=1, cols=1, style='Table Grid')
            table.cell(0, 0).text = 'Ведомость материалов для проведения расчетов и оценки негативного воздействия'
            doc.add_paragraph()
            table = doc.add_table(rows=len(materials), cols=3, style='Table Grid')
            table.autofit = False
            table.allow_autofit = False
            table.columns[0].width = Inches(1.0)
            table.rows[0].cells[0].width = Inches(1.0)
            table.columns[0].width = Inches(1.0)
            table.rows[0].cells[1].width = Inches(1.0)
            x = 0
            y = 0
            for row in table.rows:
                for cell in row.cells:
                    cell.text = str(materials[x][y])
                    y += 1
                x += 1
                y = 0
            for row in table.rows:
                for cell in row.cells:
                    paragraphs = cell.paragraphs
                    paragraph = paragraphs[0]
                    run_obj = paragraph.runs
                    run = run_obj[0]
                    font = run.font
                    font.name = 'Arial'
                    font.size = Pt(10)
            saves = QFileDialog.getSaveFileName(self,
                                                self.tr("Save file"), fn,
                                                self.tr("Doc files (*.docx *.doc)"))[0]
            if saves:
                try:
                    doc.save(saves)
                except PermissionError:
                    QMessageBox.critical(self, "Ошибка ", "Файл не может быть сохранен, так как открыт в другой программе.", QMessageBox.Ok)
        except NameError:
            QMessageBox.critical(self, "Ошибка ", "Не выбран файл", QMessageBox.Ok)

    def init_UI(self):
        self.setWindowTitle("Анализ ресурсной ведомости")
        self.UI.name_pd.setPlaceholderText('Введите название объекта')
        self.UI.chifr.setPlaceholderText('Введите шифр ПД')
        self.UI.action_5.triggered.connect(self.showDialog)
        self.UI.download_v.clicked.connect(self.Save_file)

app = QApplication([])
application = resourses_sheet()
application.show()

sys.exit(app.exec_())
input('Press ENTER to exit')