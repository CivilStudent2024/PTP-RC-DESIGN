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



class ControlParameterWindow(QWidget) :
    window_closed = pyqtSignal()
    def __init__(self,filename) :
        super().__init__()
        self.setWindowTitle("Auto Rc")
        self.setFixedSize(550,550) #550 650
        self.setStyleSheet('background-color : None')

        self.filename = filename
        self.filepath = self.filename

        try :
            self.Head_And_Title_df1 = pd.read_excel(self.filepath,sheet_name='Head_Title')
        except :
            self.Head_And_Title_df1 = pd.DataFrame({'Project' : ['-'] , 'Floor Layer' : ['-'] , 'Engineer' : ['-'], 'Date' : ['-']})
        try :
            self.Control_Parameter_df = pd.read_excel(self.filepath,sheet_name='Control_Parameter')
        except :
            self.Control_Parameter_df = pd.DataFrame({'จำนวนจุดต่อ' : [''] , 'จำนวนแผ่นพื้น' : [''] , 'จำนวนคาน' : [''], 'จำนวนประเภทหน้าตัดคาน' : [''], 'จำนวนน้ำหนักบรรทุกแบบจุด' : [''],'จำนวนน้ำหนักบรรทุกแบบกระจาย' : [''] })
















        OutFrame = QFrame(self) ; OutFrame.setFrameShape(QFrame.Shape.Box); OutFrame.setFrameShadow(QFrame.Shadow.Sunken) ; OutFrame.setLineWidth(3) 
        OutFrame.setFixedSize(480,510);OutFrame.setStyleSheet('border:5px solid purple') ; OutFrame.move(35,20)
        frame = QFrame(self) ; frame.setFrameShape(QFrame.Shape.Box); frame.setFrameShadow(QFrame.Shadow.Sunken) ; frame.setLineWidth(3) ;frame.setFixedSize(450,250);frame.setStyleSheet('background:rgba(128, 0, 128,120)') ; frame.move(50,200)
        for i, name in enumerate(['Project Title :','Floor Layer :','Engineer :','Date :']) :
            lb = QLabel(name,self) ; lb.setStyleSheet('color : purple') ; font = QFont() ; font.setPointSize(15) ; lb.setFont(font) ; lb.move(70,30*(i+1))

        #สร้างLabel จาก dataframe Head_And_Title
        for Column in range(self.Head_And_Title_df1.shape[1]) :
            lb = QLabel(f'{self.Head_And_Title_df1.iloc[0][Column]}',self) ; lb.setStyleSheet('color : green'); font = QFont() ; font.setPointSize(15) ; lb.setFont(font) ; lb.move(210,30*(Column+1))

        
        for i , name in enumerate(['จำนวนจุดต่อ','จำนวนแผ่นพื้น','จำนวนคาน','จำนวนประเภทหน้าตัดคาน','จำนวนน้ำหนักบรรทุกแบบจุด','จำนวนน้ำหนักบรรทุกแบบกระจาย']) :
            lb = QLabel(name,self) ; lb.setStyleSheet('color : yellow') ; font = QFont() ; font.setPointSize(12) ; lb.setFont(font) ; lb.move(250,190+(35*(i+1)))
            globals()[f'InfoControl{i}'] = QLineEdit(self); globals()[f'InfoControl{i}'].setText(f'{self.Control_Parameter_df.loc[0,name]}') ; globals()[f'InfoControl{i}'] ; globals()[f'InfoControl{i}'].setAlignment(Qt.AlignmentFlag.AlignHCenter);globals()[f'InfoControl{i}'].setFixedSize(140,30);globals()[f'InfoControl{i}'].move(100,190+(35*(i+1)))
            
        
        self.lbTitle = QLabel('ตั้งค่า Control Parameter',self) ;  font = QFont() ; font.setPointSize(20) ; self.lbTitle.setFont(font) ;self.lbTitle.setStyleSheet('color : yellow') ; self.lbTitle.move(150,160)
        #save button
        self.savebtn = QPushButton('บันทึก',self) ; self.savebtn.setFixedSize(QSize(150,40)) ;self.savebtn.setStyleSheet('background:black') ; self.savebtn.move(320,470)
        #signal
        self.savebtn.clicked.connect(self.SaveData)



    def SaveData(self) :
        self.Control_Parameter_df.loc[0,'จำนวนจุดต่อ'] = int(globals()[f'InfoControl{0}'].text())
        self.Control_Parameter_df.loc[0,'จำนวนแผ่นพื้น'] = int(globals()[f'InfoControl{1}'].text())
        self.Control_Parameter_df.loc[0,'จำนวนคาน'] = int(globals()[f'InfoControl{2}'].text())
        self.Control_Parameter_df.loc[0,'จำนวนประเภทหน้าตัดคาน'] = int(globals()[f'InfoControl{3}'].text())
        self.Control_Parameter_df.loc[0,'จำนวนน้ำหนักบรรทุกแบบจุด'] = int(globals()[f'InfoControl{4}'].text())
        self.Control_Parameter_df.loc[0,'จำนวนน้ำหนักบรรทุกแบบกระจาย'] = int(globals()[f'InfoControl{5}'].text())
        save_excel_sheet(self.Control_Parameter_df, self.filepath, sheetname = 'Control_Parameter', index=False)
        QMessageBox.information(self,'Slab Data Saved','บันทึกข้อมูลเรียบร้อยแล้ว')

    def keyPressEvent(self, event):
        """ ตรวจจับการกดปุ่ม Enter """
        if event.key() == 16777220:  # 16777220 คือรหัสคีย์สำหรับปุ่ม Enter
            self.SaveData()
        elif event.key() == 16777237: # Code for "Down Arrow"
            self.focusNextChild()
        elif event.key() == 16777235:  # Code for "Up Arrow"
            self.focusPreviousChild()
        elif event.key() == 16777236:  # Code for "Up Arrow"
            self.focusNext()
        elif event.key() == 16777236:  # Code for "Up Arrow"
            self.focusPrevious()
        super().keyPressEvent(event)

    def closeEvent(self, event) :
        reply = QMessageBox.question(self,"ปิดหน้าต่าง","คุณแน่ใจไหมว่าต้องการออกจากหน้าต่างนี้",QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes :
            self.window_closed.emit()
            event.accept()
        else:
            event.ignore()






if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ControlParameterWindow(filename= r'C:\ProjectVenv\.venv\ProjectHouse\budha.xlsx')
    window.show()
    sys.exit(app.exec())