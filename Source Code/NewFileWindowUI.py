from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
import pandas as pd
import os
import sys


def save_excel_sheet(df, filepath, sheetname, index=False):
    # Create file if it does not exist
    if not os.path.exists(filepath):
        df.to_excel(filepath, sheet_name=sheetname, index=index)

    # Otherwise, add a sheet. Overwrite if there exists one with the same name.
    else:
        with pd.ExcelWriter(filepath, engine='openpyxl', if_sheet_exists='replace', mode='a') as writer:
            df.to_excel(writer, sheet_name=sheetname, index=index)

class TitleWindow(QWidget) :
    window_closed = pyqtSignal()
    def __init__(self,filename) :
        super().__init__()
        self.setWindowTitle("Head And Title")
        self.setFixedSize(550,300) #550 650
        self.setStyleSheet('background-color : None')
        self.filename = filename
        self.filepath = self.filename
        try :
            self.Head_And_Title_df1 = pd.read_excel(self.filepath,sheet_name='Head_Title')
        except :
            self.Head_And_Title_df1 = pd.DataFrame({'Project :' : [''] , 'Floor Layer :' : [''] , 'Engineer :' : [''], 'Date :' : ['']})




        frame = QFrame(self) ; frame.setFrameShape(QFrame.Shape.Box); frame.setFrameShadow(QFrame.Shadow.Sunken) ; frame.setLineWidth(3) ;frame.setFixedSize(450,200);frame.setStyleSheet('background:rgba(128, 0, 128,120)') ; frame.move(50,30)
        for i, name in enumerate(['Project :','Floor Layer :','Engineer :','Date :']) :
            lb = QLabel(name,self) ; lb.setStyleSheet('color : Yellow') ; font = QFont() ; font.setPointSize(15) ; lb.setFont(font) ; lb.move(70,20+35*(i+1))
            globals()[f'InfoTitle{i}'] = QLineEdit(self); globals()[f'InfoTitle{i}'].setText(f'{self.Head_And_Title_df1.loc[0,name]}') ; globals()[f'InfoTitle{i}'] ; globals()[f'InfoTitle{i}'].setAlignment(Qt.AlignmentFlag.AlignHCenter);globals()[f'InfoTitle{i}'].setFixedSize(250,30);globals()[f'InfoTitle{i}'].move(200,20+(35*(i+1)))
        #สร้างLabel จาก dataframe Head_And_Title

        #save button
        self.savebtn = QPushButton('บันทึก',self) ; self.savebtn.setFixedSize(QSize(150,40)) ;self.savebtn.setStyleSheet('background:black') ; self.savebtn.move(320,250)
        #signal
        self.savebtn.clicked.connect(self.SaveData)

    def SaveData(self) :
        self.Head_And_Title_df1.loc[0,'Project :'] = globals()[f'InfoTitle{0}'].text()
        self.Head_And_Title_df1.loc[0,'Floor Layer :'] = globals()[f'InfoTitle{1}'].text()
        self.Head_And_Title_df1.loc[0,'Engineer :'] = globals()[f'InfoTitle{2}'].text()
        self.Head_And_Title_df1.loc[0,'Date :'] = globals()[f'InfoTitle{3}'].text()
        save_excel_sheet(self.Head_And_Title_df1, self.filename, sheetname = 'Head_Title', index=False)
        QMessageBox.information(self,'Slab Data Saved','บันทึกข้อมูลเรียบร้อยแล้ว')
    
    def keyPressEvent(self, event):
        """ ตรวจจับการกดปุ่ม Enter """
        if event.key() == 16777220:  # 16777220 คือรหัสคีย์สำหรับปุ่ม Enter
            self.SaveData()

    def closeEvent(self, event) :
        reply = QMessageBox.question(self,"ปิดหน้าต่าง","คุณแน่ใจไหมว่าต้องการออกจากหน้าต่างนี้",QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes :
            self.window_closed.emit()
            event.accept()
        else:
            event.ignore()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TitleWindow(filename = r'C:\Users\bugpi\AppData\Local\Programs\Microsoft VS Code\Project venv\.venv\Project\Arr1.xlsx')
    window.show()
    sys.exit(app.exec())