#!/Users/Hendri/venv/py27/bin/python
#############################################################################
##
## Copyright (C) 2016 
## Programmer: Hendri L Tobing
##
## This file may be used under the terms of the GNU General Public
## License version 2.0 as published by the Free Software Foundation
## and appearing in the file LICENSE.GPL included in the packaging of
## this file.
##
## This file is provided AS IS with NO WARRANTY OF ANY KIND, INCLUDING THE
## WARRANTY OF DESIGN, MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE.
##
#############################################################################

# Import required modules
import sys
import datetime
from PySide import QtGui, QtCore
from PySide.QtCore import Qt
from PySide.QtGui import *

from xlrd import open_workbook
from xlrd import open_workbook

#"Horizontal bar"
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.colors import white, PCMYKColor
from reportlab.platypus import Table, TableStyle, SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


#paper size
width, height = landscape(A4)
styles = getSampleStyleSheet()
styleN = styles["BodyText"]
styleN.alignment = TA_LEFT
styleBH = styles["Normal"]
styleBH.alignment = TA_CENTER


#Change work location
import os
saveloc = os.path.dirname(os.path.realpath(__file__))
os.chdir(saveloc)
#Time
now = datetime.datetime.now()
waktu = now.strftime("%Y-%m-%d")

values = []

#include new font

ttfFile='Designosaur-Regular.ttf'
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(TTFont("Designosaur-Regular", ttfFile))

class MainWindow(QMainWindow):
    """ Our Main Window class
    """
    def __init__(self, fileName=None):
        """ Constructor Function
        """
        QMainWindow.__init__(self)
        self.setWindowTitle("Test Report Generator App")
        self.setWindowIcon(QIcon('appicon.png'))
        self.setGeometry(100, 100, 1000, 600)
        #self.list.setWindowTitle('List of Text File')
        #self.list.setMinimumSize(100, 600)

        self.fileName = None
        self.filters = "Text files (*.txt)"

    def SetupComponents(self):
        """ Function to setup status bar, central widget, menu
            bar, tool bar, also dock widget
        """
        self.myStatusBar = QStatusBar()
        self.setStatusBar(self.myStatusBar)
        self.myStatusBar.showMessage('App is Ready', 10000)
        self.CreateActions()
        self.CreateToolBar()

        #Create Toolbar
        self.mainToolBar.addAction(self.openAction)
        self.mainToolBar.addAction(self.generateAction)
        self.mainToolBar.addAction(self.aboutAction)
        self.mainToolBar.addAction(self.exitAction)

        #dock
        self.tabbed=QTabWidget()
        self.setCentralWidget(self.tabbed)
        self.tab1=QTableWidget(0, 0)
        self.tabbed.addTab(self.tab1, "&Hasil Test")

    def setIcon(self):
        """ Function to set Icon
        """
        appIcon = QIcon('logo.png')
        self.setWindowIcon(appIcon)

    # Slots called when the menu actions are triggered
    def showDialog(self):
        fname, _ = QtGui.QFileDialog.getOpenFileName(self, 'Open file',
                    saveloc)

        #self.lbl0.setText(fname)
        #self.lbl0.adjustSize()
        #print(fname)

        #read Excel
        if fname.endswith('.xls') or fname.endswith('.xlsx'):
            global values
            values = []
            siswa = -1
            wb = open_workbook(fname)
            for s in wb.sheets():
                for row in range(s.nrows):
                    col_value = []
                    for col in range(s.ncols):
                        value  = (s.cell(row,col).value)
                        if col == 0:
                            try : value = str(int(value))
                            except : pass
                        col_value.append(value)
                    values.append(col_value)
                    siswa = siswa + 1

            #global values
            print siswa

            try:
                self.tab1.setColumnCount(len(values[0]))
                self.tab1.setRowCount(len(values))

                for i in range(len(values)):
                    for j in range(len(values[0])):
                        try:
                            item = QTableWidgetItem(str(values[i][j]))
                            self.tab1.setItem(i, j, item)
                        except ValueError:
                            pass
                        except IndexError:
                            pass

            except IndexError:
                pass
                self.tab1.setColumnCount(0)
                self.tab1.setRowCount(0)
                self.myStatusBar.showMessage('Index Error', 10000)
            except ValueError:
                pass

        else:
            self.lbl0.setText("Tipe file tidak betul!")
            self.lbl0.adjustSize()
            self.myStatusBar.showMessage('File Error', 10000)

    def onShowQuestion(self):
        try:
            NN1 = str(values[0][5])
            NN2 = str(values[0][6])
            NN3 = str(values[0][7])
            NN4 = str(values[0][8])
            NN5 = str(values[0][9])
            NN6 = str(values[0][10])
            NN7 = str(values[0][11])
            NN8 = str(values[0][12])
            NN9 = str(values[0][13])
            NN10 = str(values[0][14])
            NN11 = str(values[0][15])
            NN12 = str(values[0][16])
            NN13 = str(values[0][17])
            NN14 = str(values[0][18])
            NN15 = str(values[0][19])
            NN16 = str(values[0][20])
            NN17 = str(values[0][21])
            #---------Print Algorithm----------#
            value = 0
            while (value < (len(values))):
                #Ntotal = int(values[value][15]*100)
                print values[value]
                if value == 0:
                    print "zero val"
                    pass
                elif value > 0:
                    nama = str(values[value][1])
                    outlet = str(values[value][2])
                    kelas = str(int(values[value][3]))
                    sekolah = str(values[value][4])
                    N1 = (values[value][5]*100)
                    N2 = (values[value][6]*100)
                    N3 = (values[value][7]*100)
                    N4 = (values[value][8]*100)
                    N5 = (values[value][9]*100)
                    N6 = (values[value][10]*100)
                    N7 = (values[value][11]*100)
                    N8 = (values[value][12]*100)
                    N9 = (values[value][13]*100)
                    N10 = (values[value][14]*100)
                    N11 = (values[value][15]*100)
                    N12 = (values[value][16]*100)
                    N13 = (values[value][17]*100)
                    N14 = (values[value][18]*100)
                    N15 = (values[value][19]*100)
                    N16 = (values[value][20]*100)
                    N17 = (values[value][21]*100)
                    #Ntotal
                    print N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11, N12
                    print kelas
                    x = values[value][22]/10
                    x = "{0:.2f}".format(x)
                    NTotal = str(x)
                    print NTotal
                    filepdf = outlet+'_'+nama+'_'+waktu+'.pdf'
                    c = canvas.Canvas(filepdf, pagesize=landscape(A4))
                    print filepdf


                    def generate_page1(c):
                        #Image----------------------------------------------------------------------
                        im1 = Image.open("left.jpg")
                        c.drawInlineImage(im1, 10, 15, width=546/2.3, height=1290/2.3)

                        im2 = Image.open("shinken.jpg")
                        c.drawInlineImage(im2, 259, height-183, width=2434/4.4, height=725/4.4)

                        im3 = Image.open("skor.jpg")
                        c.drawInlineImage(im3, 259, height-251, width=2335/4.2, height=271/4.2)

                        im4 = Image.open("diagram.jpg")
                        c.drawInlineImage(im4, 259, 15, width=2339/4.2, height=1367/4.2)

                        #grafik1----------------------------------------------------------------------

                        from reportlab.graphics.shapes import Drawing
                        from reportlab.graphics.charts.barcharts import HorizontalBarChart

                        drawing = Drawing(500, 250)
                        data = [
                               (N1,N2,N3,N4,N5,N6,N7,N8,N9,N10,N11,N12,N13,N14,N15,N16,N17)
                               ]
                        bc = HorizontalBarChart()
                        bc.x = 100
                        bc.y = 100
                        bc.height = 250
                        bc.width = 300
                        bc.data = data
                        bc.strokeColor = None
                        bc.fillColor = None
#                        bc.bars[0].fillColor = PCMYKColor(92,47,0,33,alpha=95)
#                        bc.bars[0].strokeColor = PCMYKColor(92,47,0,33,alpha=95)
#                        bc.bars[0].fillColor = PCMYKColor(92,32,0,33,alpha=95)
#                        bc.bars[0].strokeColor = PCMYKColor(92,32,0,33,alpha=95)
                        bc.bars[0].fillColor = PCMYKColor(92,16,0,33,alpha=95)
                        bc.bars[0].strokeColor = PCMYKColor(92,16,0,33,alpha=95)
#                        bc.bars[0].fillColor = PCMYKColor(92,7,0,33,alpha=95)
#                        bc.bars[0].strokeColor = PCMYKColor(92,7,0,33,alpha=95)
                        bc.barWidth             = 15
                        bc.valueAxis.valueMin   = 0
                        bc.valueAxis.valueMax   = 100
                        bc.valueAxis.valueStep  = 10
                        bc.valueAxis.visibleAxis = False
                        bc.valueAxis.visibleGrid = False
                        bc.valueAxis.visibleTicks = False
                        bc.valueAxis.forceZero = True
                        bc.valueAxis.visibleLabels  = 0

                        bc.categoryAxis.visibleGrid         = False
                        bc.categoryAxis.visibleTicks        = False # hidding the ticks remove the label
                        bc.categoryAxis.tickLeft            = 0    # a workaround is to set the tick length
                        bc.categoryAxis.tickRight           = 0    # to zero.
                        bc.categoryAxis.strokeWidth         = 0.25
                        bc.categoryAxis.labelAxisMode       ='low'
                        bc.categoryAxis.labels.textAnchor   ='end'
                        bc.categoryAxis.labels.angle        = 0
                        bc.categoryAxis.labels.fontName     = 'Designosaur-Regular'
                        #bc.categoryAxis.labels.fontColor    = PCMYKColor(0,65,100,0,alpha=90)
                        bc.categoryAxis.labels.boxAnchor    = 'e'
                        bc.categoryAxis.labels.dx           = -5
                        bc.categoryAxis.labels.dy           = 0
                        bc.categoryAxis.labels.angle        = 0
                        bc.categoryAxis.reverseDirection    = 1
                        bc.categoryAxis.joinAxisMode        ='left'
                        bc.categoryAxis.categoryNames = [NN1,NN2,NN3,NN4,NN5,NN6,NN7,NN8,NN9,NN10,NN11,NN12,NN13,NN14,NN15,NN16,NN17]

                        bc.barLabels.fontName      = 'Designosaur-Regular'
                        bc.barLabels.fontSize      = 10
                        bc.barLabels.angle         = 0
                        bc.barLabelFormat          = "%.00f%%"
                        bc.barLabels.boxAnchor     ='w'
                        bc.barLabels.boxFillColor  = None
                        bc.barLabels.boxStrokeColor= None
                        bc.barLabels.dx            = 10
                        bc.barLabels.dy            = 0
                        #bc.barLabels.dy            = -1
                        bc.barLabels.boxTarget     = 'hi'


                        drawing.add(bc)
                        drawing.wrapOn(c, width, height)
                        drawing.drawOn(c, 350, height - 650)


                        #Table---------------------------------------------------------------------

                        styleBH = styles["Normal"]
                        styleBH.alignment = TA_CENTER
                        styleBH.fontSize = 12
                        styleBH.fontName = 'Designosaur-Regular'
                        #styleBH.textColor = PCMYKColor(92,7,0,33,alpha=100)
                        isi1 = Paragraph(nama, styleBH)
                        isitabel1 = [[isi1]]
                        table1 = Table(isitabel1, colWidths=[150])
                        table1.wrapOn(c, width, height)
                        table1.drawOn(c, 54, height-352)

                        isi2 = Paragraph(kelas, styleBH)
                        isitabel2 = [[isi2]]
                        table2 = Table(isitabel2, colWidths=[150])
                        table2.wrapOn(c, width, height)
                        table2.drawOn(c, 54, height-386)

                        isi3 = Paragraph(sekolah, styleBH)
                        isitabel3 = [[isi3]]
                        table3 = Table(isitabel3, colWidths=[150])
                        table3.wrapOn(c, width, height)
                        table3.drawOn(c, 54, height-421)

                        isi4 = Paragraph(outlet, styleBH)
                        isitabel4 = [[isi4]]
                        table4 = Table(isitabel4, colWidths=[150])
                        table4.wrapOn(c, width, height)
                        table4.drawOn(c, 54, height-455)


                        #Text----------------------------------------------------------------------

                        #nilai total
                        c.setFillColor(white)
                        c.setStrokeColor(white)
                        if (NTotal == '10.0'):
                            c.setFont("Designosaur-Regular", 35)
                            c.drawString(352, height-230, NTotal +"/")
                            c.setFont("Designosaur-Regular", 20)
                            c.drawString(395, height-235, "10")
                        else:
                            c.setFont("Designosaur-Regular", 35)
                            c.drawString(338, height-230, NTotal +"/")
                            c.setFont("Designosaur-Regular", 20)
                            c.drawString(415, height-235, "10")

                        #Yang paling tidak dikuasai siswa
                        c.setFont("Designosaur-Regular", 9)
                        c.drawString(576, height-205, "Topik yang paling tidak dikuasai peserta Try-Out II :")
                        c.drawString(576, height-216, "1. Mean, Median, dan Modus")
                        c.drawString(576, height-227, "2. Debit Air")
                        c.drawString(576, height-238, "3. Bangun Ruang")


                    def generate_page2(c):
                        im2 = Image.open("backside.jpg")
                        c.drawInlineImage(im2, 30, 12, width=3326/4.3, height=2424/4.3)

                    generate_page1(c)
                    c.showPage()

                    generate_page2(c)
                    c.save()
                value = value + 1
                #value2 = str(value)
                print value
                #self.myStatusBar.showMessage('Mencetak '+ str(value) +'/'+str(len(values)), 10000)
            QtGui.QMessageBox.information(self, "Information", "Proses membuat laporan selesai!")
            now2 = datetime.datetime.now()
            print now2
            #---------Print Algorithm Ends----------#
        except NameError:
            print(sys.exc_info()[1])
            QtGui.QMessageBox.information(self, "Information", "Pastikan Anda telah membuka file yang benar!!")
        except Exception:
            print(sys.exc_info()[1])
            QtGui.QMessageBox.information(self, "Information", "Pastikan Anda telah membuka file yang benar!!!")




    def exitFile(self):
        self.close()

    def aboutHelp(self):
        QMessageBox.about(self, "About Report", "An Application that help generate and print PDF Report")

    def CreateActions(self):
        """ Function to create actions for menus
        """
        self.openAction = QAction( QIcon('open.png'), 'O&pen',
                                   self, shortcut=QKeySequence.Open,
                                   statusTip="Open an existing file",
                                   triggered=self.showDialog)
        self.generateAction = QAction( QIcon('generate.png'), 'P&review',
                                  self, shortcut='Ctrl+P',
                                  statusTip="Generate PDF Report(s)",
                                  triggered=self.onShowQuestion)
        self.aboutAction = QAction( QIcon('about.png'), 'A&bout',
                                    self, statusTip="Displays info about the Application",
                                    triggered=self.aboutHelp)
        self.exitAction = QAction( QIcon('exit.png'), 'E&xit',
                                  self, shortcut="Ctrl+Q",
                                  statusTip="Exit the Application",
                                  triggered=self.exitFile)



    def CreateToolBar(self):
        self.mainToolBar = self.addToolBar('Main')




if __name__ == '__main__':
# Exception Handling
    try:
        myApp = QApplication(sys.argv)
        mainWindow = MainWindow()
        mainWindow.SetupComponents()
        mainWindow.setIcon()
        mainWindow.show()
        myApp.exec_()
        sys.exit(0)
    except NameError:
        print("Name Error:", sys.exc_info()[1])
    except SystemExit:
        print("Closing Window...")
    except Exception:
        print(sys.exc_info()[1])
