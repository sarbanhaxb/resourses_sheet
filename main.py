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

        electrodes = 0 #электроды
        paints = 0 #ЛКМ
        geosynthetic_nt = 0 #геосинтетическая мембрана нетканная
        geogrids_nt = 0 #георешетка нетканная
        thermal = 0 #теплоизоляционный материал
        bituminous = 0 #битумно-резиновая мастика
        sands = 0 #песок
        geomembranes = 0 #геомембрана
        slabs = 0 #дорожные плиты
        semens = 0 #семена
        fertilized = 0 #удобрения
        wire = 0 #провода

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
                    sands += float(row['amount'])*1.6*1000
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
                    semens += float(row['amount'])*1
                elif str(row['unit']).lower() == 'т':
                    semens += float(row['amount'])/1000
            """расчет удобрений"""
            if str(row['code']).lower().find('16.3.02.') != -1 and (str(row['name']).lower().find('удобрен') == 0 or str(row['name']).lower().find('селитра') == 0):
                if str(row['unit']).lower() == 'кг':
                    fertilized += float(row['amount'])*1
            """расчет проводов"""
            if str(row['code']).lower().find('21.2.01.01') != -1 and str(row['name']).lower().find('провод') == 0:
                """Группа 21.2.01.01. Провода изолированные для воздушных линий электропередач"""
                if str(row['unit']).lower() == '1000 м':
                    if str(row['name']).lower().find('сип-1') != -1:
                        if str(row['name']).lower().find('1x16+1x25-0,6/1') != -1:
                            wire += float(row['amount'])*1000*0.14
                        elif str(row['name']).lower().find('3x16+1x25-0,6/1') != -1:
                            wire += float(row['amount'])*1000*0.271
                        elif str(row['name']).lower().find('3x25+1x35-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.381
                        elif str(row['name']).lower().find('3x35+1x50-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.524
                        elif str(row['name']).lower().find('3x50+1x50-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.727
                        elif str(row['name']).lower().find('3x50+1x70-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.727
                        elif str(row['name']).lower().find('3x70+1x70-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.923
                        elif str(row['name']).lower().find('3x70+1x70-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.994
                        elif str(row['name']).lower().find('3x95+1x70-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 1.195
                        elif str(row['name']).lower().find('3x95+1x95-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 1.266
                        elif str(row['name']).lower().find('3x120+1x95-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 1.466
                        elif str(row['name']).lower().find('3x150+1x95-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 1.741
                        elif str(row['name']).lower().find('3x185+1x95-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 2.33
                        elif str(row['name']).lower().find('3x240+1x95-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 2.895
                    elif str(row['name']).lower().find('сип-2') != -1:
                        if str(row['name']).lower().find('3х16+1х25-0,6/1') != -1:
                            wire += float(row['amount'])*1000*0.315
                        elif str(row['name']).lower().find('3x16+1x54,6-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.410
                        elif str(row['name']).lower().find('3x16+1x54,6+2x16-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.579
                        elif str(row['name']).lower().find('3x16+1x54,6+2x35-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.812
                        elif str(row['name']).lower().find('3x25+1x35-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.422
                        elif str(row['name']).lower().find('3х25+54,6-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.526
                        elif str(row['name']).lower().find('3х25+54,6+2х16-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.664
                        elif str(row['name']).lower().find('3x25+1x54,6+2x35-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.896
                        elif str(row['name']).lower().find('3x35+1x50-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.568
                        elif str(row['name']).lower().find('3x35+1x50+2x16-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.708
                        elif str(row['name']).lower().find('3х35+54,6') != -1:
                            wire += float(row['amount']) * 1000 * 0.735
                        elif str(row['name']).lower().find('3x35+1x54,6+2x35-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.62 #не точно
                        elif str(row['name']).lower().find('3x50+1x50-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.727
                        elif str(row['name']).lower().find('3x50+1x50+2x16-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.866
                        elif str(row['name']).lower().find('3х50+54,6') != -1:
                            wire += float(row['amount']) * 1000 * 0.776
                        elif str(row['name']).lower().find('3x50+1x70-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.8
                        elif str(row['name']).lower().find('3x50+1x70+2x35-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.97
                        elif str(row['name']).lower().find('3х70+54,6-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.99
                        elif str(row['name']).lower().find('3х70+54,6+2х16-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 1.128
                        elif str(row['name']).lower().find('3х70+70') != -1:
                            wire += float(row['amount']) * 1000 * 1.012
                        elif str(row['name']).lower().find('3x70+1x70+2x35-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 1.180
                        elif str(row['name']).lower().find('3x70+1x95-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 1.093
                        elif str(row['name']).lower().find('3x95+1x70-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 1.170
                        elif str(row['name']).lower().find('3x95+1x95-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 1.266
                        elif str(row['name']).lower().find('3x120+1x95-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 1.549
                        elif str(row['name']).lower().find('3x150+1x95-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 1.799
                        elif str(row['name']).lower().find('3x185+1x95-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 2.146
                        elif str(row['name']).lower().find('3x240+1x95-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 2.65
                        elif str(row['name']).lower().find('4x16+1x25-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.386
                        elif str(row['name']).lower().find('4x25+1x35-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.525
                    elif str(row['name']).lower().find('сип-3') != -1:
                        if str(row['name']).lower().find('1x35-20') != -1:
                            wire += float(row['amount']) * 1000 * 0.165
                        elif str(row['name']).lower().find('1x35-35') != -1:
                            wire += float(row['amount']) * 1000 * 0.209
                        elif str(row['name']).lower().find('1х50-20') != -1:
                            wire += float(row['amount']) * 1000 * 0.215
                        elif str(row['name']).lower().find('1x50-35') != -1:
                            wire += float(row['amount']) * 1000 * 0.263
                        elif str(row['name']).lower().find('1х70-20') != -1:
                            wire += float(row['amount']) * 1000 * 0.282
                        elif str(row['name']).lower().find('1x70-35') != -1:
                            wire += float(row['amount']) * 1000 * 0.445
                        elif str(row['name']).lower().find('1х95-20') != -1:
                            wire += float(row['amount']) * 1000 * 0.364
                        elif str(row['name']).lower().find('1x95-35') != -1:
                            wire += float(row['amount']) * 1000 * 0.421
                        elif str(row['name']).lower().find('1х120-20') != -1:
                            wire += float(row['amount']) * 1000 * 0.445
                        elif str(row['name']).lower().find('1x120-35') != -1:
                            wire += float(row['amount']) * 1000 * 0.518
                        elif str(row['name']).lower().find('1х150-20') != -1:
                            wire += float(row['amount']) * 1000 * 0.54
                        elif str(row['name']).lower().find('1x150-35') != -1:
                            wire += float(row['amount']) * 1000 * 0.618
                        elif str(row['name']).lower().find('1x185-20') != -1:
                            wire += float(row['amount']) * 1000 * 0.722
                        elif str(row['name']).lower().find('1x185-35') != -1:
                            wire += float(row['amount']) * 1000 * 0.808
                        elif str(row['name']).lower().find('1x240-20') != -1:
                            wire += float(row['amount']) * 1000 * 0.95
                        elif str(row['name']).lower().find('1x240-35') != -1:
                            wire += float(row['amount']) * 1000 * 1.045
                    elif str(row['name']).lower().find('сип-4') != -1:
                        if str(row['name']).lower().find('2x10-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.092
                        elif str(row['name']).lower().find('2х16') != -1:
                            wire += float(row['amount']) * 1000 * 0.136
                        elif str(row['name']).lower().find('2х25-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.194
                        elif str(row['name']).lower().find('2x50-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.355
                        elif str(row['name']).lower().find('4х16-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.261
                        elif str(row['name']).lower().find('4х25-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.388
                        elif str(row['name']).lower().find('4x35-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.48
                        elif str(row['name']).lower().find('4x50-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.709
                        elif str(row['name']).lower().find('4x70-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 0.981
                        elif str(row['name']).lower().find('4x95-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 1.295
                        elif str(row['name']).lower().find('4x120-0,6/1') != -1:
                            wire += float(row['amount']) * 1000 * 1.6
            elif str(row['code']).lower().find('21.2.01.02') != -1 and str(row['name']).lower().find('провод') == 0:
                """Группа 21.2.01.02. Провода неизолированные для воздушных линий электропередач"""
                if str(row['unit']).lower() == 'т':
                    wire += float(row['amount'])*1000
            elif str(row['code']).lower().find('21.2.03.05') != -1 and str(row['name']).lower().find('провод') == 0:
                """Группа 21.2.03.05. Провода силовые для электрических установок на напряжение до 450 в"""
                if str(row['unit']).lower() == '1000 м':
                    if str(row['name']).lower().find('апв') != -1:
                        if str(row['name']).lower().find('2-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.014 #неточно
                        elif str(row['name']).lower().find('2,5-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.014
                        elif str(row['name']).lower().find('3-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.018
                        elif str(row['name']).lower().find('4-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.021
                        elif str(row['name']).lower().find('5-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.025
                        elif str(row['name']).lower().find('8-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.04
                        elif str(row['name']).lower().find('10-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.043
                        elif str(row['name']).lower().find('16-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.072
                        elif str(row['name']).lower().find('25-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.105
                        elif str(row['name']).lower().find('35-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.137
                        elif str(row['name']).lower().find('50-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.203
                        elif str(row['name']).lower().find('70-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.269
                        elif str(row['name']).lower().find('95-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.375
                        elif str(row['name']).lower().find('120-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.456
                    elif str(row['name']).lower().find('аппв') != -1:
                        if str(row['name']).lower().find('2x2-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.031 #не точно
                        elif str(row['name']).lower().find('2x2,5-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.031
                        elif str(row['name']).lower().find('2x3-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.031 #не точно
                        elif str(row['name']).lower().find('2x4-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.043
                        elif str(row['name']).lower().find('2x5-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.043 #не точно
                        elif str(row['name']).lower().find('2x6-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.057  # не точно
                        elif str(row['name']).lower().find('3x2-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.048  # не точно
                        elif str(row['name']).lower().find('3x2,5-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.048  # не точно
                        elif str(row['name']).lower().find('3x3-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.048  # не точно
                        elif str(row['name']).lower().find('3x4-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.065
                        elif str(row['name']).lower().find('3x5-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.074
                        elif str(row['name']).lower().find('3x6-450') != -1:
                            wire += float(row['amount']) * 1000 * 0.085

                        #https://fgisrf.ru/fsscm/21.2.03.05/ закончил на строке 21.2.03.05-0032	Провод силовой установочный АППВ 3x6-450
                    elif str(row['name']).lower().find('ВПИСАТЬ КЛЮЧЕВОЕ СЛОВО') != -1:
                        if str(row['name']).lower().find('РАЗМЕРЫ КАБЕЛЯ') != -1:
                            wire += float(row['amount']) * 1000 * 0.0
                        elif str(row['name']).lower().find('РАЗМЕРЫ КАБЕЛЯ') != -1:
                            wire += float(row['amount']) * 1000 * 0.0
                        elif str(row['name']).lower().find('РАЗМЕРЫ КАБЕЛЯ') != -1:
                            wire += float(row['amount']) * 1000 * 0.0
                        elif str(row['name']).lower().find('РАЗМЕРЫ КАБЕЛЯ') != -1:
                            wire += float(row['amount']) * 1000 * 0.0
                        elif str(row['name']).lower().find('РАЗМЕРЫ КАБЕЛЯ') != -1:
                            wire += float(row['amount']) * 1000 * 0.0
                        elif str(row['name']).lower().find('РАЗМЕРЫ КАБЕЛЯ') != -1:
                            wire += float(row['amount']) * 1000 * 0.0



        global materials
        materials = [['Наименование основных строительных конструкций, изделий и материалов', 'Единица измерения', 'Всего']]
        electrodes = ['Электроды', 'кг', round(electrodes*1000, 2)]
        paints = ['Лакокрасочные материалы', 'кг', round(paints, 2)]
        geosynthetic_nt = ['Нетканое геополотно', 'м2', round(geosynthetic_nt, 2)]
        geogrids_nt = ['Нетканая георешетка', 'м2', round(geogrids_nt, 2)]
        thermal = ['Теплоизоляционный материал', 'м3', round(thermal, 2)]
        bituminous = ['Битумно-резиновая мастика', 'кг', round(bituminous, 2)]
        sands = ['Песок', 'т', round(sands, 2)/1000]
        geomembranes = ['Геомембрана', 'м2', round(geomembranes, 2)]
        slabs = ['Плиты дорожные', 'т', round(slabs, 2)]
        semens = ['Семена трав', "кг", round(semens, 2)]
        fertilized = ['Удобрения', "кг", round(fertilized, 2)]
        wire = ['Провод', 'кг', round(wire, 2)]

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
        if fertilized[2] != 0:
            materials.append(fertilized)
        if wire[2] != 0:
            materials.append(wire)

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