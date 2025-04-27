import sys
import os
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
import pandas as pd


def save_excel_sheet(df, filepath, sheetname, index=False):
    # Create file if it does not exist
    if not os.path.exists(filepath):
        df.to_excel(filepath, sheet_name=sheetname, index=index)

    # Otherwise, add a sheet. Overwrite if there exists one with the same name.
    else:
        with pd.ExcelWriter(filepath, engine='openpyxl', if_sheet_exists='replace', mode='a') as writer:
            df.to_excel(writer, sheet_name=sheetname, index=index)



class PointLoadWindow(QWidget):
    window_closed = pyqtSignal()
    def __init__(self,filename):
        super().__init__()
        # สร้าง QGraphicsView
        self.view = QGraphicsView(self)
        self.view.setStyleSheet('background-color : wireframe')
        self.view.setFixedSize(QSize(550,440))
        self.frame = QFrame(self) ; self.frame.setFrameShape(QFrame.Shape.Box); self.frame.setFrameShadow(QFrame.Shadow.Sunken)
        self.frame.setLineWidth(3) ;self.frame.setFixedSize(400,480);self.frame.setStyleSheet('background:None') ; self.frame.move(50,30)
        self.view.move(510,50)
        self.view.setDragMode(self.view.DragMode.ScrollHandDrag)
        self.setWindowTitle("PointLoad Data")
        self.setFixedSize(1100, 550)
        # สร้าง QGraphicsScene
        self.scene = QGraphicsScene(self)
        # กำหนดขนาดของ Scene และสร้างกรอบ
        self.scene.setSceneRect(0, 0, 1500, 1500) #550 440
        # สร้างกรอบให้กับ Scene
        # กำหนด QGraphicsView ให้ใช้ Scene
        self.view.setScene(self.scene)
        #lable widget
        #grid widget
        self.grid_widget = QWidget(self)
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidget(self.grid_widget)
        self.scroll_area.setFixedSize(QSize(380,440))
        self.scroll_area.move(60,50)

        self.filename = filename
        self.filepath = self.filename

        self.Node_df = pd.read_excel(self.filepath,sheet_name='Node_Data')
        self.Control_Parameter_df = pd.read_excel(self.filepath,sheet_name='Control_Parameter')

        try : 
            self.PointLoad_df = pd.read_excel(self.filepath,sheet_name='PointLoad_Data')
        except :
            self.PointLoad_df = pd.DataFrame({'หมายเลขจุดต่อ' : [] , 'สถานะจุดต่อ' : [] , 'ขนาดน้ำหนักบรรทุก (T/m)' : [],'Section No.' : []})
            for i in range(self.Control_Parameter_df.loc[0,'จำนวนจุดต่อ']) :
                self.Node_df.loc[i,'สถานะจุดต่อ'] = 'N'



        for i in range(int(self.Control_Parameter_df.loc[0,'จำนวนจุดต่อ'])) :
            self.PointLoad_df.loc[i,'หมายเลขจุดต่อ'] = int(i+1)

        self.PointLoad_df = self.PointLoad_df.fillna('')


        if max(self.Node_df['Node No. (int)']) >= 13 :
            self.grid_widget.setFixedSize(QSize(380,440+(50*max(self.Node_df['Node No. (int)']))))
        else :
            self.grid_widget.setFixedSize(QSize(380,440))

        for i,name in enumerate(["จุดต่อที่","Status","Load (T/m)","Section No."]) :
            lb = QLabel(name,self.grid_widget) ; lb.setStyleSheet('color : yellow') ; font = QFont() ; font.setPointSize(9) ; lb.setFont(font) ; lb.move(25+(i*95),30) ; lb.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        for i,val in enumerate(self.PointLoad_df['หมายเลขจุดต่อ']) :
            globals()[f'Node-No_{i}'] = QLineEdit(str(self.PointLoad_df.loc[i,'หมายเลขจุดต่อ']),self.grid_widget) ; globals()[f'Node-No_{i}'].setFixedSize(60,20);globals()[f'Node-No_{i}'].move(10,60+((i)*30)) ;globals()[f'Node-No_{i}'].setStyleSheet('color : yellow');globals()[f'Node-No_{i}'].setAlignment(Qt.AlignmentFlag.AlignCenter)
            globals()[f'Node-Status_{i}'] = QLineEdit(str(self.PointLoad_df.loc[i,'สถานะจุดต่อ']),self.grid_widget); globals()[f'Node-Status_{i}'].setFixedSize(60,20);globals()[f'Node-Status_{i}'].move(105,60+((i)*30)) ; globals()[f'Node-Status_{i}'].setStyleSheet('color : green');globals()[f'Node-Status_{i}'].setAlignment(Qt.AlignmentFlag.AlignCenter)    
            globals()[f'PointLoad{i}'] = QLineEdit(str(self.PointLoad_df.loc[i,'ขนาดน้ำหนักบรรทุก (T/m)']),self.grid_widget) ; globals()[f'PointLoad{i}'].setFixedSize(60,20);globals()[f'PointLoad{i}'].move(210,60+((i)*30)) ;globals()[f'PointLoad{i}'].setStyleSheet('color : green');globals()[f'PointLoad{i}'].setAlignment(Qt.AlignmentFlag.AlignCenter)
            globals()[f'section_node{i}'] = QLineEdit(str(self.PointLoad_df.loc[i,'Section No.']),self.grid_widget); globals()[f'section_node{i}'].setFixedSize(60,20);globals()[f'section_node{i}'].move(305,60+((i)*30)) ; globals()[f'section_node{i}'].setStyleSheet('color : green');globals()[f'section_node{i}'].setAlignment(Qt.AlignmentFlag.AlignCenter)
            
        self.btnsubmit = QPushButton("บันทึก",self)
        self.btnsubmit.setFixedSize(QSize(100,50))
        self.btnsubmit.move(950,500)



        #signal
        self.btnsubmit.clicked.connect(self.submit)
        

    def submit(self) :
        for i,value in enumerate(self.PointLoad_df['หมายเลขจุดต่อ']) :
            self.PointLoad_df.loc[i,'สถานะจุดต่อ'] = str(globals()[f'Node-Status_{i}'].text())
            self.Node_df.loc[i,'สถานะจุดต่อ'] = str(globals()[f'Node-Status_{i}'].text())
            self.PointLoad_df.loc[i,'ขนาดน้ำหนักบรรทุก (T/m)'] = str(globals()[f'PointLoad{i}'].text())
            self.PointLoad_df.loc[i,'Section No.'] = str(globals()[f'section_node{i}'].text())
            
            
        self.view.centerOn(200, 750)
        self.scene.clear()        
        self.PltNode()
        self.DimDraw()
        save_excel_sheet(self.PointLoad_df, self.filepath, sheetname='PointLoad_Data', index=False)
        save_excel_sheet(self.Node_df, self.filepath, sheetname='Node_Data', index=False)
        
        QMessageBox.information(self,'info','บันทึกข้อมูลเรียบร้อยแล้ว')
    



    def PltNode(self) :

        for i,value in enumerate(self.Node_df['Node No. (int)']) :
            print(i)
            x = float(self.Node_df.loc[i,'พิกัด X (m)(.2float)']) * 20 ;y = float(self.Node_df.loc[i,'พิกัด Y (m)  (.2float)']) *20
            Nodetag = QGraphicsTextItem(str(self.Node_df.loc[i,'Node No. (int)'])) ; font=QFont("Terminal"); font.setPixelSize(8)
            Nodetag.setFont(font) ; Nodetag.setDefaultTextColor(QColor('gold'))
            Nodetag.setPos(300+x,700-y) #285 700
            self.scene.addItem(Nodetag)

            
            if float(self.PointLoad_df.loc[i,'ขนาดน้ำหนักบรรทุก (T/m)']) != 0.0  :
                loadtag = QGraphicsTextItem(f"{self.PointLoad_df.loc[i,'ขนาดน้ำหนักบรรทุก (T/m)']} T") ; font=QFont("Terminal"); font.setPixelSize(8) ;loadtag.setFont(font) ;  loadtag.setDefaultTextColor(QColor('red'))                
                loadtag.setPos(300+x,700-y-15)
                self.scene.addItem(loadtag)

            #ตัวล่อแนวทางการวาดสี่เหลี่ยม จุดเริ่มต้น(0,0) คือพิกัด (300,700)
            if self.Node_df.loc[i,'สถานะจุดต่อ'] == 'Y' :
                recttest = QGraphicsRectItem(300+x-3, 700-y-3, 6, 6) ; recttest.setBrush(QBrush(QColor(255, 0, 0)))
            elif self.Node_df.loc[i,'สถานะจุดต่อ'] == 'N' :
                recttest = QGraphicsRectItem(300+x-2, 700-y-2, 4, 4) ; recttest.setPen(QPen(QColor(255,0,0)))
            elif self.Node_df.loc[i,'สถานะจุดต่อ'] == 'E' :
                recttest = QGraphicsRectItem(300+x-3, 700-y-3, 6, 6) ; recttest.setBrush(QBrush(QColor(0, 255, 0)))
            else : recttest = QGraphicsRectItem(300+x-3, 700-y-3, 6, 6) ; recttest.setBrush(QBrush(QColor(255, 0, 0)))
            #recttest = QGraphicsRectItem(297, 697, 6, 6) ; recttest.setBrush(QBrush(QColor(255, 0, 0))) #ปิดไว้
            self.scene.addItem(recttest)



    def DimDraw(self) :
        x = self.Node_df['พิกัด Y (m)  (.2float)'].tolist()
        x = list(set(x))
        x.sort()
        y = self.Node_df['พิกัด X (m)(.2float)'].tolist()
        y = list(set(y))
        y.sort()
        print(y)
        
        L = 1000
        for i , val in enumerate(x) :
            if i != 0 :
                X = round(val,3)*20
                dx2x1 = round(x[i]-x[i-1],3) *20
                offset = -min(x) * 20 
                Hdim = QGraphicsLineItem(230-offset, 700-X,230-offset, 700-X+dx2x1 ) ; Hdim.setPen(QPen(Qt.GlobalColor.darkGreen))
                Dimtext = QGraphicsTextItem(str(dx2x1/20)) ; font=QFont("Terminal") ; font.setPixelSize(8) ; Dimtext.setFont(font);Dimtext.setRotation(-90)
                Dimtext.setPos(210-offset,(20+700-X+700-X+(dx2x1))/2) ; Dimtext.setDefaultTextColor(QColor('gold'))
                self.scene.addItem(Hdim)
                self.scene.addItem(Dimtext)
                vline1 = QGraphicsLineItem(225-offset, 700-X,235-offset, 700-X) ; vline1.setPen(QPen(Qt.GlobalColor.darkGreen))
                vline2 = QGraphicsLineItem(225-offset, 700-X+dx2x1 ,235-offset, 700-X+dx2x1 ) ; vline2.setPen(QPen(Qt.GlobalColor.darkGreen))
                self.scene.addItem(vline1) ;self.scene.addItem(vline2)

        for i , val in enumerate(y) :
            if i != 0 :
                X = round(val,3)*20
                dy2y1 = round(y[i]-y[i-1],3) *20
                offset = -min(y) * 20 
                Hdim = QGraphicsLineItem(300+X, 1000-230+offset,300+X-dy2y1,1000-230+offset) ; Hdim.setPen(QPen(Qt.GlobalColor.darkGreen))
                Dimtext = QGraphicsTextItem(str(round((dy2y1/20),3))) ; font.setPixelSize(8) ; Dimtext.setFont(font)
                Dimtext.setPos(-12+300+X-((dy2y1)/2),1000-225+offset) ; Dimtext.setDefaultTextColor(QColor('gold'))
                self.scene.addItem(Hdim)
                self.scene.addItem(Dimtext)
                vline1 = QGraphicsLineItem(300+X, 1000-225+offset , 300+X, 1000-235+offset ) ; vline1.setPen(QPen(Qt.GlobalColor.darkGreen))
                vline2 = QGraphicsLineItem(300+X-dy2y1, 1000-225+offset ,300+X-dy2y1, 1000-235+offset ) ; vline2.setPen(QPen(Qt.GlobalColor.darkGreen))
                self.scene.addItem(vline1) ;self.scene.addItem(vline2)

  

    def closeEvent(self, event) :
        reply = QMessageBox.question(self,"ปิดหน้าต่าง","คุณแน่ใจไหมว่าต้องการออกจากหน้าต่างนี้",QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes :
            self.window_closed.emit()
            event.accept()
        else:
            event.ignore()

    def keyPressEvent(self, event):
        """ ตรวจจับการกดปุ่ม Enter """
        if event.key() == 16777220:  # 16777220 คือรหัสคีย์สำหรับปุ่ม Enter
            self.submit()
        elif event.key() == 16777237: # Code for "Down Arrow"
            self.focusNextChild()
        elif event.key() == 16777235:  # Code for "Up Arrow"
            self.focusPreviousChild()
        elif event.key() == 16777236:  # Code for "Up Arrow"
            self.focusNext()
        elif event.key() == 16777236:  # Code for "Up Arrow"
            self.focusPrevious()
        elif event.key() == Qt.Key.Key_X:
            self.view.scale(1/1.1,1/1.1)
        elif event.key() == Qt.Key.Key_Z:
            self.view.scale(1.1,1.1)
        super().keyPressEvent(event) 


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PointLoadWindow(filename = r'C:\ProjectVenv\.venv\ProjectHouse\Exhouse.xlsx')
    window.show()
    sys.exit(app.exec())