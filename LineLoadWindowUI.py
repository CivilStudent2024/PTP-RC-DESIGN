import sys
import os
import math
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



class LineLoadWindowui(QWidget):
    window_closed = pyqtSignal()
    def __init__(self,filename):
        super().__init__()
        # สร้าง QGraphicsView
        self.view = QGraphicsView(self)
        self.view.setStyleSheet('background-color : wireframe')
        self.view.setFixedSize(QSize(550,440))
        self.frame = QFrame(self) ; self.frame.setFrameShape(QFrame.Shape.Box); self.frame.setFrameShadow(QFrame.Shadow.Sunken)
        self.frame.setLineWidth(3) ;self.frame.setFixedSize(400,480);self.frame.setStyleSheet('background:None') ; self.frame.move(50,30)
        self.view.setDragMode(self.view.DragMode.ScrollHandDrag)
        self.view.move(510,50)
        self.setWindowTitle("LineLoad Data")
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
        self.Beam_df = pd.read_excel(self.filepath,sheet_name='Beam_Data')
 
        try : 
            self.LineLoad_df = pd.read_excel(self.filepath,sheet_name='Lineload_Data')
        except :
            self.LineLoad_df = pd.DataFrame({'DATA No.' : [] , 'NodeList1' : [] , 'NodeList2' : [],'Load(T/M)' : []})

        for i in range(int(self.Control_Parameter_df.loc[0,'จำนวนน้ำหนักบรรทุกแบบกระจาย'])) :
            self.LineLoad_df.loc[i,'DATA No.'] = int(i+1)
        self.LineLoad_df = self.LineLoad_df.fillna(0)

        if self.LineLoad_df.empty == False:
            if max(self.LineLoad_df['DATA No.']) >= 13 :
                self.grid_widget.setFixedSize(QSize(380,440+(50*int(max(self.LineLoad_df['DATA No.'])))))
            else :
                self.grid_widget.setFixedSize(QSize(380,440))
        else : self.grid_widget.setFixedSize(QSize(380,440))

        for i,name in enumerate(["DATA No.","จุดเริ่มต้น","จุดสุดท้าย","Load (T/M)"]) :
            lb = QLabel(name,self.grid_widget) ; lb.setStyleSheet('color : yellow') ; font = QFont() ; font.setPointSize(9) ; lb.setFont(font) ; lb.move(25+(i*95),30) ; lb.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        if self.LineLoad_df.empty == False:
            for i,val in enumerate(self.LineLoad_df['DATA No.']) :
                globals()[f'Node-No_{i}'] = QLineEdit(str(self.LineLoad_df.loc[i,'DATA No.']),self.grid_widget) ; globals()[f'Node-No_{i}'].setFixedSize(60,20);globals()[f'Node-No_{i}'].move(10,60+((i)*30)) ;globals()[f'Node-No_{i}'].setStyleSheet('color : yellow');globals()[f'Node-No_{i}'].setAlignment(Qt.AlignmentFlag.AlignCenter)
                globals()[f'Node-Status_{i}'] = QLineEdit(str(self.LineLoad_df.loc[i,'NodeList1']),self.grid_widget); globals()[f'Node-Status_{i}'].setFixedSize(60,20);globals()[f'Node-Status_{i}'].move(105,60+((i)*30)) ; globals()[f'Node-Status_{i}'].setStyleSheet('color : green');globals()[f'Node-Status_{i}'].setAlignment(Qt.AlignmentFlag.AlignCenter)    
                globals()[f'PointLoad{i}'] = QLineEdit(str(self.LineLoad_df.loc[i,'NodeList2']),self.grid_widget) ; globals()[f'PointLoad{i}'].setFixedSize(60,20);globals()[f'PointLoad{i}'].move(210,60+((i)*30)) ;globals()[f'PointLoad{i}'].setStyleSheet('color : green');globals()[f'PointLoad{i}'].setAlignment(Qt.AlignmentFlag.AlignCenter)
                globals()[f'section_node{i}'] = QLineEdit(str(self.LineLoad_df.loc[i,'Load(T/M)']),self.grid_widget); globals()[f'section_node{i}'].setFixedSize(60,20);globals()[f'section_node{i}'].move(305,60+((i)*30)) ; globals()[f'section_node{i}'].setStyleSheet('color : green');globals()[f'section_node{i}'].setAlignment(Qt.AlignmentFlag.AlignCenter)
                
        self.btnsubmit = QPushButton("บันทึก",self)
        self.btnsubmit.setFixedSize(QSize(100,50))
        self.btnsubmit.move(950,500)



        #signal
        self.btnsubmit.clicked.connect(self.submit)
        

    def submit(self) :
        if self.LineLoad_df.empty == False :
            for i,value in enumerate(self.LineLoad_df['DATA No.']) :
                self.LineLoad_df.loc[i,'DATA No.'] = str(globals()[f'Node-Status_{i}'].text())
                self.LineLoad_df.loc[i,'NodeList1'] = str(globals()[f'Node-Status_{i}'].text())
                self.LineLoad_df.loc[i,'NodeList2'] = str(globals()[f'PointLoad{i}'].text())
                self.LineLoad_df.loc[i,'Load(T/M)'] = str(globals()[f'section_node{i}'].text())
                

                Node1check = float(self.LineLoad_df.loc[i,'NodeList1'])
                Node2check = float(self.LineLoad_df.loc[i,'NodeList2'])
                for Rw,Node in enumerate(self.Node_df['Node No. (int)']) :
                    if int(Node) == Node1check :
                        self.LineLoad_df.loc[i,'Node1x'] =  float(self.Node_df.loc[Rw,'พิกัด X (m)(.2float)'])
                        self.LineLoad_df.loc[i,'Node1y'] =  float(self.Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])
                    elif int(Node) == Node2check :
                        self.LineLoad_df.loc[i,'Node2x'] =  float(self.Node_df.loc[Rw,'พิกัด X (m)(.2float)'])
                        self.LineLoad_df.loc[i,'Node2y'] =  float(self.Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])
                if self.LineLoad_df.loc[i,'Node1x'] != self.LineLoad_df.loc[i,'Node2x'] and self.LineLoad_df.loc[i,'Node1y'] == self.LineLoad_df.loc[i,'Node2y'] :
                    self.LineLoad_df.loc[i,'direction'] = 'X'
                elif self.LineLoad_df.loc[i,'Node1x'] == self.LineLoad_df.loc[i,'Node2x'] and self.LineLoad_df.loc[i,'Node1y'] != self.LineLoad_df.loc[i,'Node2y'] :
                    self.LineLoad_df.loc[i,'direction'] = 'Y'




                
            self.view.centerOn(200, 750)
            self.scene.clear()  
            self.DimDraw()
            self.BeamDraw()   
            self.LinloadDraw()
            self.PltNode()
            
            save_excel_sheet(self.LineLoad_df, self.filepath, sheetname='Lineload_Data', index=False)
            
            QMessageBox.information(self,'info','บันทึกข้อมูลเรียบร้อยแล้ว')
        else : print('ไม่มีข้อมูล')



    def PltNode(self) :

        for i,value in enumerate(self.Node_df['Node No. (int)']) :
            print(i)
            x = float(self.Node_df.loc[i,'พิกัด X (m)(.2float)']) * 20 ;y = float(self.Node_df.loc[i,'พิกัด Y (m)  (.2float)']) *20
            Nodetag = QGraphicsTextItem(str(self.Node_df.loc[i,'Node No. (int)'])) ; font=QFont("Terminal"); font.setPixelSize(8)
            Nodetag.setFont(font) ; Nodetag.setDefaultTextColor(QColor('gold'))
            Nodetag.setPos(300+x,700-y) #285 700
            self.scene.addItem(Nodetag)


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

    def BeamDraw(self) :
        for i,value in enumerate(self.Beam_df['ทิศทางการวาง']) :
            xOrg = 300
            yOrg = 700
            if value == 'X' :
                pos1 = [self.Beam_df.loc[i,'Node1x'] *20 ,self.Beam_df.loc[i,'Node1y'] *20] 
                pos2 = [self.Beam_df.loc[i,'Node2x'] *20 ,self.Beam_df.loc[i,'Node2y'] *20] 
                Hbeam = QGraphicsRectItem(xOrg+pos1[0]-2,yOrg-pos1[1]-2, (pos2[0]-pos1[0]), 4) ; Hbeam.setBrush(QBrush(QColor('None'))) ; Hbeam.setPen(QPen(QColor('grey')))
                self.scene.addItem(Hbeam)



            elif value == 'Y' :
                pos1 = [self.Beam_df.loc[i,'Node1x'] *20 ,self.Beam_df.loc[i,'Node1y'] *20] 
                pos2 = [self.Beam_df.loc[i,'Node2x'] *20 ,self.Beam_df.loc[i,'Node2y'] *20] 
                Vbeam = QGraphicsRectItem(xOrg+pos1[0]-2,yOrg-pos1[1]-2, 4, (-pos2[1]+pos1[1])) ; Vbeam.setBrush(QBrush(QColor('None'))) ; Vbeam.setPen(QPen(QColor('grey')))
                self.scene.addItem(Vbeam)
  
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

    def LinloadDraw(self) :
        for i,value in enumerate(self.LineLoad_df['direction']) :
            xOrg = 300
            yOrg = 700
            pos1 = [self.LineLoad_df.loc[i,'Node1x'] *20 ,self.LineLoad_df.loc[i,'Node1y'] *20] 
            pos2 = [self.LineLoad_df.loc[i,'Node2x'] *20 ,self.LineLoad_df.loc[i,'Node2y'] *20] 
            if value == 'X' :

                lenload = (pos2[0]-pos1[0])/3
                for j in range(math.floor(lenload)) :
                    loadline = QGraphicsLineItem(xOrg+pos1[0]-1+(3*j), yOrg-pos1[1]+1,xOrg+pos1[0]-1+((3*j)+1), yOrg-pos1[1]-1 ); loadline.setPen(QPen(Qt.GlobalColor.red))
                    self.scene.addItem(loadline)            
                load_tag = QGraphicsTextItem(f"{self.LineLoad_df.loc[i,'Load(T/M)']} T/M") ; font=QFont("Terminal"); font.setPixelSize(7) ; load_tag.setFont(font) ; load_tag.setDefaultTextColor(QColor('red'))
                load_tag.setPos(-10+xOrg+pos1[0]+(pos2[0]-pos1[0])/2,-20+yOrg-pos1[1])
                self.scene.addItem(load_tag) 
            else :
                lenload = (pos2[1]-pos1[1])/3
                for j in range(math.floor(lenload)) :
                    loadline = QGraphicsLineItem(xOrg+pos1[0]+1, yOrg-pos1[1]+1-(3*j),xOrg+pos1[0]-1, yOrg-pos1[1]-1-((3*j)+1) ); loadline.setPen(QPen(Qt.GlobalColor.red))
                    self.scene.addItem(loadline)  
                load_tag = QGraphicsTextItem(f"{self.LineLoad_df.loc[i,'Load(T/M)']} T/M") ; font=QFont("Terminal"); font.setPixelSize(7) ; load_tag.setFont(font) ; load_tag.setDefaultTextColor(QColor('red'))
                load_tag.setPos(xOrg+pos1[0]-20,20+yOrg-pos1[1]-(pos2[1]-pos1[1])/2) ; load_tag.setRotation(-90)
                self.scene.addItem(load_tag)
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
        elif event.key() == Qt.Key.Key_X:
            self.view.scale(1/1.1,1/1.1)
        elif event.key() == Qt.Key.Key_Z:
            self.view.scale(1.1,1.1)
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
    window = LineLoadWindowui(filename = r'C:\ProjectVenv\.venv\ProjectHouse\Ex6_3.xlsx')
    window.show()
    sys.exit(app.exec())
    