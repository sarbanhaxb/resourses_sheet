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
        """расчет электродов"""
        for index, row in df.iterrows():
            if str(row['name']).lower().find('электрод') == 0:
                if str(row['unit']).lower() == 'кг':
                    electrodes += float(row['amount']) / 1000
                elif str(row['unit']).lower() == 'т':
                    electrodes += float(row['amount'])
        global materials
        materials = [['Наименование основных строительных конструкций, изделий и материалов', 'Единица измерения', 'Всего']]
        electrodes = ['Электроды', 'кг', round(electrodes*1000, 3)]
        materials.append(electrodes)

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
        table.columns[0].width = Inches(1.5)
        table.rows[0].cells[0].width = Inches(3.0)
        table.columns[0].width = Inches(1.5)
        table.rows[0].cells[1].width = Inches(1)
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
            doc.save(saves)

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