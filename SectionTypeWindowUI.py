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




class SectionTypeWindow(QWidget) :
    window_closed = pyqtSignal()
    def __init__(self,filename,DL) :
        super().__init__()
        self.setWindowTitle("ประเภทหน้าตัดคาน")
        self.setFixedSize(675,600) #675 600
        self.setStyleSheet('background-color : None')
        self.filename = filename
        self.filepath = self.filename        
        self.Head_And_Title_df1 = pd.read_excel(self.filepath,sheet_name='Head_Title')
        self.Control_Parameter_df = pd.read_excel(self.filepath,sheet_name='Control_Parameter')
        self.DL = float(DL)

        try : 
            self.SectionType_df = pd.read_excel(self.filepath,sheet_name='SectionType')
        except :
            self.SectionType_df = pd.DataFrame({'หมายเลขหน้าตัดคาน' : [],'ความกว้าง (m)' : [],'ความลึก (m)' : [], 'ระยะหุ้มเหล็กเสริมบน (m)' : [], 'ระยะหุ้มเหล็กเสริมล่าง (m)' : [],'SelfWeight' : []})

        for i in range(int(self.Control_Parameter_df.loc[0,'จำนวนประเภทหน้าตัดคาน'])) :
            self.SectionType_df.loc[i,'หมายเลขหน้าตัดคาน'] = str(i+1)

        OutFrame = QFrame(self) ; OutFrame.setFrameShape(QFrame.Shape.Box); OutFrame.setFrameShadow(QFrame.Shadow.Sunken) ; OutFrame.setLineWidth(3) 
        OutFrame.setFixedSize(630,560);OutFrame.setStyleSheet('border:5px solid purple') ; OutFrame.move(25,15)
        frame = QFrame(self) ; frame.setFrameShape(QFrame.Shape.Box); frame.setFrameShadow(QFrame.Shadow.Sunken) ; frame.setLineWidth(3) ;frame.setFixedSize(580,340);frame.setStyleSheet('background:rgba(128, 0, 128,120)') ; frame.move(50,160)
        for i, name in enumerate(['Project Title :','Floor Layer :','Engineer :','Date :']) :
            lb = QLabel(name,self) ; lb.setStyleSheet('color : purple') ; font = QFont() ; font.setPointSize(15) ; lb.setFont(font) ; lb.move(70,30*(i+1))

        #สร้างLabel จาก dataframe Head_And_Title
        for Column in range(self.Head_And_Title_df1.shape[1]) :
            lb = QLabel(f'{self.Head_And_Title_df1.iloc[0][Column]}',self) ; lb.setStyleSheet('color : green'); font = QFont() ; font.setPointSize(15) ; lb.setFont(font) ; lb.move(210,30*(Column+1))
        
        for i , name in enumerate(['จำนวนประเภทหน้าตัดคาน','หมายเลขหน้าตัดคาน']) :
            lb = QLabel(name,self) ; lb.setStyleSheet('color : white') ; font = QFont() ; font.setPointSize(14) ; lb.setFont(font) ; lb.move(130,190+(35*(i+1)))
            if name == 'จำนวนประเภทหน้าตัดคาน' :
                noslabData =  str(self.Control_Parameter_df.loc[0,'จำนวนประเภทหน้าตัดคาน'])
                noslabDataEntry = QLineEdit(noslabData,self);noslabDataEntry.setReadOnly(True) ;noslabDataEntry.setStyleSheet('color : yellow'); noslabDataEntry.setAlignment(Qt.AlignmentFlag.AlignHCenter);noslabDataEntry.setFixedSize(100,30);noslabDataEntry.move(350,190+(35*(i+1)))
            elif name == 'หมายเลขหน้าตัดคาน' :
                self.SectionCombo = QComboBox(self) ; self.SectionCombo.addItems(str(i) for i in range(1,int(self.Control_Parameter_df.loc[0,'จำนวนประเภทหน้าตัดคาน'])+1));self.SectionCombo.setFixedSize(100,30);self.SectionCombo.move(350,190+(35*(i+1)))

        
        for i , name in enumerate(['ความกว้าง (m)','ความลึก (m)','ระยะหุ้มเหล็กเสริมบน (m)','ระยะหุ้มเหล็กเสริมล่าง (m)']) :
            lb = QLabel(name,self) ; lb.setStyleSheet('color : yellow') ; font = QFont() ; font.setPointSize(12) ; lb.setFont(font) ; lb.move(150,300+(30*(i+1))) ; lb.setAlignment(Qt.AlignmentFlag.AlignRight)
            globals()[f'info{i}'] = QLineEdit(self); globals()[f'info{i}'].setText(f'{str(self.SectionType_df.iloc[self.SectionCombo.currentIndex()][i+1])}')  ; globals()[f'info{i}'].setAlignment(Qt.AlignmentFlag.AlignHCenter);globals()[f'info{i}'].setFixedSize(180,30);globals()[f'info{i}'].move(350,300+(30*(i+1)))
            


        #save button
        self.savebtn = QPushButton('บันทึก',self) ; self.savebtn.setFixedSize(QSize(160,50)) ; self.savebtn.move(450,510)
        #signal

        self.SectionCombo.currentIndexChanged.connect(self.UpdateValue)
        self.savebtn.clicked.connect(self.SaveData)



        
    def SaveData(self) :
        IndexInfo = self.SectionCombo.currentIndex()
        self.SectionType_df.loc[IndexInfo,'ความกว้าง (m)'] = float(globals()[f'info{0}'].text())
        self.SectionType_df.loc[IndexInfo,'ความลึก (m)'] = float(globals()[f'info{1}'].text())
        self.SectionType_df.loc[IndexInfo,'ระยะหุ้มเหล็กเสริมบน (m)'] = float(globals()[f'info{2}'].text())
        self.SectionType_df.loc[IndexInfo,'ระยะหุ้มเหล็กเสริมล่าง (m)'] = float(globals()[f'info{3}'].text())
        self.SectionType_df.loc[IndexInfo,'SelfWeight'] = 2.4*self.SectionType_df.loc[IndexInfo,'ความกว้าง (m)'] * self.SectionType_df.loc[IndexInfo,'ความลึก (m)']*self.DL           
        save_excel_sheet(self.SectionType_df, self.filepath, sheetname = 'SectionType', index=False)
        QMessageBox.information(self,'Section Data Saved','บันทึกข้อมูลเรียบร้อยแล้ว')
                
    def UpdateValue(self) :
        IndexInfo = self.SectionCombo.currentIndex()
        globals()[f'info{0}'].setText(str(self.SectionType_df.loc[IndexInfo,'ความกว้าง (m)']))
        globals()[f'info{1}'].setText(str(self.SectionType_df.loc[IndexInfo,'ความลึก (m)']))
        globals()[f'info{2}'].setText(str(self.SectionType_df.loc[IndexInfo,'ระยะหุ้มเหล็กเสริมบน (m)']))
        globals()[f'info{3}'].setText(str(self.SectionType_df.loc[IndexInfo,'ระยะหุ้มเหล็กเสริมล่าง (m)']))
      
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
    window = SectionTypeWindow(filename = r'C:\ProjectVenv\.venv\ProjectHouse\Exhouse - Copy.xlsx')
    window.show()

    sys.exit(app.exec())