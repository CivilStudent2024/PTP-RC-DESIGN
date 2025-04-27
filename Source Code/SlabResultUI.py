import pyqtgraph as pg
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import QFont,QColor,QPixmap,QImage,QPainter
import numpy as np
import sys
import pandas as pd

class SlabResultWindow(QWidget) :
    window_closed = pyqtSignal()
    def __init__(self,filepath,DL,LL) :
        super().__init__()
        self.setWindowTitle("Auto Rc")
        self.setFixedSize(1300,700) #595 842
        self.setStyleSheet('background-color : None')
        self.central_widget = QWidget(self)
        self.layout = QVBoxLayout(self.central_widget)
        self.central_widget.move(25,90)

        self.filepath = filepath
        self.table_layout = QVBoxLayout()
        self.tableWidget = QTableWidget(self)

        self.tableWidget.setFixedSize(1200,560) ; self.tableWidget.move(50,75)
        self.tableWidget.setColumnWidth(0, 140)
        self.tableWidget.verticalHeader().setVisible(False)

        self.tableWidget.setStyleSheet("QHeaderView::section { background-color: purple; color: yellow; }")
        for i in range(0,8) :
            self.tableWidget.setColumnWidth(i, 50) 

        self.load_excel_data(1)
        self.table_layout.addWidget(self.tableWidget)
        
        title_lb = QLabel(f'ผลการวิเคราะห์และออกแบบพื้น Load Combination {DL}DL + {LL}LL',self) ; titlefont = QFont() ; titlefont.setPointSize(20),title_lb.setFont(titlefont) ;title_lb.setStyleSheet('color : white');  title_lb.move(50,15)


    def UpdateValue(self) :

        nobeam = self.nobeam.currentIndex() + 1 


        self.plot_area.clear()
        self.load_excel_data(nobeam)

    def load_excel_data(self,NO):
        df = pd.read_excel(self.filepath,sheet_name=f'Slab_Result')
        df.fillna('',inplace=True)

        self.tableWidget.setRowCount(df.shape[0])
        self.tableWidget.setColumnCount(df.shape[1])
        for i in range(1,8) : 
            self.tableWidget.setColumnWidth(i,100)
        self.tableWidget.setHorizontalHeaderLabels(df[['หมายเลขพื้น','SlabType','Case','Xlength','Ylength','Thickness','Top Cover','Bot Cover',
                                                       'AS#1','AS#2','AS#3','AS#4','AS#5','AS#6','Use ST#1','Use ST#2','Use ST#3',
                                                       'Use ST#4','Use ST#5','Use ST#6']])

        # Insert the data into the QTableWidget
        for row in range(df.shape[0]):
            for col in range(df.shape[1]):
                if col in [1,2,3,4,5,6,7] :
                    try :        
                        item = QTableWidgetItem(f'{df.iloc[row, col]:.3f}') ; item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
                    except : item = QTableWidgetItem(f'{df.iloc[row, col]}') ; item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
                else : item = QTableWidgetItem(f'{df.iloc[row, col]}') ; item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
                self.tableWidget.setItem(row, col, item)
        self.set_column_colors()

    def set_column_colors(self):
        for col in range(self.tableWidget.columnCount()):
            for row in range(self.tableWidget.rowCount()):
                item = self.tableWidget.item(row, col)
                
                if col % 2 == 0:
                    item.setBackground(QColor('#9F5499'))
                else:
                    pass
                    item.setBackground(QColor('#484B52'))

    def closeEvent(self, event) :
        reply = QMessageBox.question(self,"ปิดหน้าต่าง","คุณแน่ใจไหมว่าต้องการออกจากหน้าต่างนี้",QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes :
            self.window_closed.emit()
            event.accept()
        else:
            event.ignore()

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    window = SlabResultWindow(filepath = r'C:\ProjectVenv\.venv\ProjectHouse\Ex6_1-Analysed.xlsx',DL= 1.4,LL=1.7)
    window.show()

    sys.exit(app.exec())
from PyQt6.QtGui import QImage, QPainter, QImageReader
from PyQt6.QtCore import Qt
import numpy as np

