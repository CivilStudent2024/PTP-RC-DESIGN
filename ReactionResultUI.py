import pyqtgraph as pg
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import QFont,QColor,QPixmap,QImage,QPainter
import numpy as np
import sys
import pandas as pd

class ReactionResultWindow(QWidget) :
    window_closed = pyqtSignal()
    def __init__(self,filepath,DL,LL) :
        super().__init__()
        self.setWindowTitle("Reaction Result")
        self.setFixedSize(270,460) #595 842
        self.setStyleSheet('background-color : None')
        self.central_widget = QWidget(self)
        self.layout = QVBoxLayout(self.central_widget)
        self.central_widget.move(25,90)

        self.filepath = filepath
        self.table_layout = QVBoxLayout()
        self.tableWidget = QTableWidget(self)

        self.tableWidget.setFixedSize(200,400) ; self.tableWidget.move(30,55)
        self.tableWidget.setColumnWidth(0, 140)
        self.tableWidget.verticalHeader().setVisible(False)

        self.tableWidget.setStyleSheet("QHeaderView::section { background-color: purple; color: yellow; }")
        for i in range(0,8) :
            self.tableWidget.setColumnWidth(i, 50) 

        self.load_excel_data(1)
        self.table_layout.addWidget(self.tableWidget)
        
        title_lb = QLabel(f'ผลลัพธ์ Reaction {DL}DL + {LL}LL',self) ; titlefont = QFont() ; titlefont.setPointSize(10),title_lb.setFont(titlefont) ;title_lb.setStyleSheet('color : white');  title_lb.move(45,15)



    def load_excel_data(self,NO):
        df = pd.read_excel(self.filepath,sheet_name=f'Reaction_result')
        df.fillna('',inplace=True)

        self.tableWidget.setRowCount(df.shape[0])
        self.tableWidget.setColumnCount(df.shape[1])
        self.tableWidget.setHorizontalHeaderLabels(['จุดต่อ','Reaction (ตัน)'])

        # Insert the data into the QTableWidget
        for row in range(df.shape[0]):
            for col in range(df.shape[1]):
                if col in [1] :
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
    window = ReactionResultWindow(filepath = r'C:\ProjectVenv\.venv\ProjectHouse\Exhouse-Analysed.xlsx',DL= 1.4,LL=1.7)
    window.show()

    sys.exit(app.exec())
from PyQt6.QtGui import QImage, QPainter, QImageReader
from PyQt6.QtCore import Qt
import numpy as np

