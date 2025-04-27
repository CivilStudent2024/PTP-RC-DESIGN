from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
import pandas as pd
import winsound
import sys
import os


def save_excel_sheet(df, filepath, sheetname, index=False):
    # Create file if it does not exist
    if not os.path.exists(filepath):
        df.to_excel(filepath, sheet_name=sheetname, index=index)

    # Otherwise, add a sheet. Overwrite if there exists one with the same name.
    else:
        with pd.ExcelWriter(filepath, engine='openpyxl', if_sheet_exists='replace', mode='a') as writer:
            df.to_excel(writer, sheet_name=sheetname, index=index)


class BeamDataWindow(QWidget) :
    window_closed = pyqtSignal()
    def __init__(self,filename) :
        super().__init__()
        self.setWindowTitle("Beam Data")
        self.setFixedSize(1200,650) #550 650
        self.setStyleSheet('background-color : None')
        self.view = QGraphicsView(self)
        self.view.setStyleSheet('background-color : wireframe')
        self.view.setFixedSize(QSize(600,550))
        self.view.setDragMode(self.view.DragMode.ScrollHandDrag)

        self.view.move(550,40)
        self.scene = QGraphicsScene(self)
        # กำหนดขนาดของ Scene และสร้างกรอบ
        self.scene.setSceneRect(0, 0, 1500, 1500) #550 440
        # สร้างกรอบให้กับ Scene

        # กำหนด QGraphicsView ให้ใช้ Scene
        self.view.setScene(self.scene)

        self.filename = filename
        self.filepath = self.filename
        try :        
            self.Grid_df = pd.read_excel(self.filepath,sheet_name='GirdLine')
            self.Grid_df = self.Grid_df.fillna('')
            self.Gx = 0 ; self.Gy = 0
            for i in self.Grid_df['Grid X'] :
                if i != '' :
                    self.Gx += 1
            for i in self.Grid_df['Grid Y'] :
                if i != '' :
                    self.Gy += 1

        except : pass
        self.Node_df = pd.read_excel(self.filepath,sheet_name='Node_Data')
        self.Head_And_Title_df1 = pd.read_excel(self.filepath,sheet_name='Head_Title')
        self.Control_Parameter_df = pd.read_excel(self.filepath,sheet_name='Control_Parameter')

        try : 
            self.Beam_df = pd.read_excel(self.filepath,sheet_name='Beam_Data')
        except :
            self.Beam_df = pd.DataFrame({'หมายเลขคาน' : [] , 'กลุ่มคาน' : [] , 'Node List1' : [],'หน้าตัดคาน' : [],'Node List2' : [],'Node1x' : [],'Node1y' : [],'Node2x' :[],'Node2y' :[]})

        for i in range(int(self.Control_Parameter_df.loc[0,'จำนวนคาน'])) :
            self.Beam_df.loc[i,'หมายเลขคาน'] = str(i+1)
        print(self.Beam_df)
        self.Beam_df = self.Beam_df.fillna(0.0)

        OutFrame = QFrame(self) ; OutFrame.setFrameShape(QFrame.Shape.Box); OutFrame.setFrameShadow(QFrame.Shadow.Sunken) ; OutFrame.setLineWidth(3) 
        OutFrame.setFixedSize(480,600);OutFrame.setStyleSheet('border:5px solid purple') ; OutFrame.move(35,20)
        frame = QFrame(self) ; frame.setFrameShape(QFrame.Shape.Box); frame.setFrameShadow(QFrame.Shadow.Sunken) ; frame.setLineWidth(3) ;frame.setFixedSize(450,350);frame.setStyleSheet('background:rgba(128, 0, 128,120)') ; frame.move(50,200)
        for i, name in enumerate(['Project Title :','Floor Layer :','Engineer :','Date :']) :
            lb = QLabel(name,self) ; lb.setStyleSheet('color : purple') ; font = QFont() ; font.setPointSize(15) ; lb.setFont(font) ; lb.move(70,30*(i+1))

        #สร้างLabel จาก dataframe Head_And_Title
        for Column in range(self.Head_And_Title_df1.shape[1]) :
            lb = QLabel(f'{self.Head_And_Title_df1.iloc[0][Column]}',self) ; lb.setStyleSheet('color : green'); font = QFont() ; font.setPointSize(15) ; lb.setFont(font) ; lb.move(210,30*(Column+1))
        
        for i , name in enumerate(['จำนวนคาน','หมายเลขคาน']) :
            lb = QLabel(name,self) ; lb.setStyleSheet('color : white') ; font = QFont() ; font.setPointSize(14) ; lb.setFont(font) ; lb.move(130,190+(35*(i+1)))
            if name == 'จำนวนคาน' :
                noslabData =  str(self.Control_Parameter_df.loc[0,'จำนวนคาน'])
                noslabDataEntry = QLineEdit(noslabData,self);noslabDataEntry.setReadOnly(True) ;noslabDataEntry.setStyleSheet('color : yellow'); noslabDataEntry.setAlignment(Qt.AlignmentFlag.AlignHCenter);noslabDataEntry.setFixedSize(140,30);noslabDataEntry.move(240,190+(35*(i+1)))
            elif name == 'หมายเลขคาน' :
                self.beamCombo = QComboBox(self) ; self.beamCombo.addItems(str(i) for i in range(1,int(self.Control_Parameter_df.loc[0,'จำนวนคาน'])+1)) ;self.beamCombo.move(240,190+(35*(i+1)))

        
        for i , name in enumerate(['จำนวนหน้าตัด','จุดเริ่มต้นคาน','หน้าตัดคาน','จุดปลายคาน']) :
            lb = QLabel(name,self) ; lb.setStyleSheet('color : yellow') ; font = QFont() ; font.setPointSize(12) ; lb.setFont(font) ; lb.move(140,300+(35*(i+1)))
            globals()[f'InfoBeam{i}'] = QLineEdit(self); globals()[f'InfoBeam{i}'].setText(f'{int(self.Beam_df.iloc[self.beamCombo.currentIndex()][i+1])}') ; globals()[f'InfoBeam{i}'] ; globals()[f'InfoBeam{i}'].setAlignment(Qt.AlignmentFlag.AlignHCenter);globals()[f'InfoBeam{i}'].setFixedSize(140,30);globals()[f'InfoBeam{i}'].move(280,300+(35*(i+1)))
            if name == 'จำนวนหน้าตัด' :
                globals()[f'InfoBeam{i}'].setText(str(1))

        
        #save button
        self.savebtn = QPushButton('บันทึก',self) ; self.savebtn.setFixedSize(QSize(150,40)) ;self.savebtn.setStyleSheet('background:black') ; self.savebtn.move(320,560)
        #signal
        self.beamCombo.currentIndexChanged.connect(self.UpdateValue)
        self.savebtn.clicked.connect(self.SaveData)


    def UpdateValue(self) :
        infobeam = self.beamCombo.currentIndex()
        globals()[f'InfoBeam{0}'].setText(str(self.Beam_df.loc[infobeam,'จำนวนหน้าตัด']))
        globals()[f'InfoBeam{1}'].setText(str(self.Beam_df.loc[infobeam,'Node List1']))
        globals()[f'InfoBeam{2}'].setText(str(self.Beam_df.loc[infobeam,'หน้าตัดคาน']))
        globals()[f'InfoBeam{3}'].setText(str(self.Beam_df.loc[infobeam,'Node List2']))



    def SaveData(self) :
        winsound.Beep(1000,500)
        Indexinfobeam = self.beamCombo.currentIndex()
        self.Beam_df.loc[Indexinfobeam,'จำนวนหน้าตัด'] = globals()[f'InfoBeam{0}'].text()
        self.Beam_df.loc[Indexinfobeam,'Node List1'] = globals()[f'InfoBeam{1}'].text()
        self.Beam_df.loc[Indexinfobeam,'หน้าตัดคาน'] = globals()[f'InfoBeam{2}'].text()
        self.Beam_df.loc[Indexinfobeam,'Node List2'] = globals()[f'InfoBeam{3}'].text()


        Node1check = int(self.Beam_df.loc[Indexinfobeam,'Node List1'])
        Node2check = int(self.Beam_df.loc[Indexinfobeam,'Node List2'])
        for Rw,Node in enumerate(self.Node_df['Node No. (int)']) :
            if int(Node) == Node1check :
                self.Beam_df.loc[Indexinfobeam,'Node1x'] =  float(self.Node_df.loc[Rw,'พิกัด X (m)(.2float)'])
                self.Beam_df.loc[Indexinfobeam,'Node1y'] =  float(self.Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])
            elif int(Node) == Node2check :
                self.Beam_df.loc[Indexinfobeam,'Node2x'] =  float(self.Node_df.loc[Rw,'พิกัด X (m)(.2float)'])
                self.Beam_df.loc[Indexinfobeam,'Node2y'] =  float(self.Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])
        if self.Beam_df.loc[Indexinfobeam,'Node1x'] != self.Beam_df.loc[Indexinfobeam,'Node2x'] and self.Beam_df.loc[Indexinfobeam,'Node1y'] == self.Beam_df.loc[Indexinfobeam,'Node2y'] :
            self.Beam_df.loc[Indexinfobeam,'ทิศทางการวาง'] = 'X'
        elif self.Beam_df.loc[Indexinfobeam,'Node1x'] == self.Beam_df.loc[Indexinfobeam,'Node2x'] and self.Beam_df.loc[Indexinfobeam,'Node1y'] != self.Beam_df.loc[Indexinfobeam,'Node2y'] :
            self.Beam_df.loc[Indexinfobeam,'ทิศทางการวาง'] = 'Y'



        
        self.view.centerOn(200, 750)
        self.scene.clear()
        self.BeamDraw()        
        self.PltNode()
        self.DimDraw()
        #self.HGridDraw()
        #self.VGridDraw()
        QMessageBox.information(self,'Slab Data Saved','บันทึกข้อมูลเรียบร้อยแล้ว')
        save_excel_sheet(self.Beam_df, self.filepath, sheetname = 'Beam_Data', index=False)    

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
    

    def VGridDraw(self) :

        for i,value in enumerate(self.Grid_df['Grid Y']):
            if i < self.Gy :
                X = value*20
                L = 400 #1000
                label = str(self.Grid_df.loc[i,'Label Y'])
                line_item = QGraphicsLineItem(300+X, 1000-100,300+X, 1000-(100+L+200))
                Circle_tag = QGraphicsEllipseItem(300+X-10,1000-100,20,20)
                Pen = QPen(QColor('green')) ; Pen.setStyle(Qt.PenStyle.DashDotLine) ;Pen.setWidth(0) ;line_item.setPen(Pen)
                #line_item.setPen(QPen(Qt.GlobalColor.green))
                Circle_tag.setPen(QPen(Qt.GlobalColor.red))
                self.scene.addItem(line_item)
                self.scene.addItem(Circle_tag)
                lbtag = QGraphicsTextItem(label) ; font=QFont("Terminal"); font.setPixelSize(8)
                lbtag.setFont(font)
                lbtag.setPos(302.5+X-10,1000-98)
                self.scene.addItem(lbtag)

        #####ถึงตรงนี้
            if  0 < i < self.Gy :
                print(self.Grid_df.loc[i,'Y2-Y1'])
                dy2y1 = self.Grid_df.loc[i,'Y2-Y1'] *20
                Hdim = QGraphicsLineItem(300+X, 1000-150,300+X-dy2y1,1000-150) ; Hdim.setPen(QPen(Qt.GlobalColor.darkBlue))
                Dimtext = QGraphicsTextItem(str(round((dy2y1/20),3))) ; font.setPixelSize(8) ; Dimtext.setFont(font)
                Dimtext.setPos(300+X+(dy2y1/20)/2,1000-170)
                self.scene.addItem(Hdim)
                self.scene.addItem(Dimtext)
    def HGridDraw(self) :

        for i,value in enumerate(self.Grid_df['Grid X']):
            if i < self.Gx  : #self.Gx จะอิงมาจาก Control Parameter NoGridX และ 20 เป็น scale factor
                print(value)
                X = value*20
                L = 1000
                line_item = QGraphicsLineItem(100, 700-X,100+L, 700-X)
                Circle_tag = QGraphicsEllipseItem(80,700-(X+10),20,20)
                Pen = QPen(QColor('green')) ; Pen.setStyle(Qt.PenStyle.DashDotLine) ;Pen.setWidth(0) ;line_item.setPen(Pen)
                #line_item.setPen(QPen(Qt.GlobalColor.green))
                Circle_tag.setPen(QPen(Qt.GlobalColor.red))
                self.scene.addItem(line_item)
                self.scene.addItem(Circle_tag)
                lbtag = QGraphicsTextItem(self.Grid_df.loc[i,'Label X']) ; font=QFont("Terminal"); font.setPixelSize(8) ; lbtag.setFont(font)
                lbtag.setPos(82.5,697.5-(X+5))
                self.scene.addItem(lbtag)
            #เขียนเส้น dimension
            if  0 < i < self.Gx :
                dx2x1 = self.Grid_df.loc[i,'X2-X1'] *20
                Hdim = QGraphicsLineItem(150, 700-X,150, 700-X+dx2x1 ) ; Hdim.setPen(QPen(Qt.GlobalColor.darkBlue))
                Dimtext = QGraphicsTextItem(str(dx2x1/20)) ; font.setPixelSize(8) ; Dimtext.setFont(font);Dimtext.setRotation(-90)
                Dimtext.setPos(130,(20+700-X+700-X+(dx2x1))/2)
                self.scene.addItem(Hdim)
                self.scene.addItem(Dimtext)
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
            #recttest = QGraphicsRectItem(297, 697, 6, 6) ; recttest.setBrush(QBrush(QColor(255, 0, 0))) #ปิดไว้
            self.scene.addItem(recttest)
    def BeamDraw(self) :
        for i,value in enumerate(self.Beam_df['ทิศทางการวาง']) :
            xOrg = 300
            yOrg = 700
            if value == 'X' :
                pos1 = [self.Beam_df.loc[i,'Node1x'] *20 ,self.Beam_df.loc[i,'Node1y'] *20] 
                pos2 = [self.Beam_df.loc[i,'Node2x'] *20 ,self.Beam_df.loc[i,'Node2y'] *20] 
                Hbeam = QGraphicsRectItem(xOrg+pos1[0]-2,yOrg-pos1[1]-2, (pos2[0]-pos1[0]), 4) ; Hbeam.setBrush(QBrush(QColor('None'))) ; Hbeam.setPen(QPen(QColor('cyan')))
                self.scene.addItem(Hbeam)
                Beamtag = QGraphicsTextItem(f"B{(self.Beam_df.loc[i,'หมายเลขคาน'])}") ; font=QFont("Terminal"); font.setPixelSize(8)
                Beamtag.setFont(font) ; Beamtag.setDefaultTextColor(QColor('cyan'))
                Beamtag.setPos(-5+xOrg+pos1[0]+(pos2[0]-pos1[0])/2,-20+yOrg-pos1[1])
                self.scene.addItem(Beamtag)

            elif value == 'Y' :
                pos1 = [self.Beam_df.loc[i,'Node1x'] *20 ,self.Beam_df.loc[i,'Node1y'] *20] 
                pos2 = [self.Beam_df.loc[i,'Node2x'] *20 ,self.Beam_df.loc[i,'Node2y'] *20] 
                Vbeam = QGraphicsRectItem(xOrg+pos1[0]-2,yOrg-pos1[1]-2, 4, (-pos2[1]+pos1[1])) ; Vbeam.setBrush(QBrush(QColor('None'))) ; Vbeam.setPen(QPen(QColor('cyan')))
                self.scene.addItem(Vbeam)
                Beamtag = QGraphicsTextItem(f"B{(self.Beam_df.loc[i,'หมายเลขคาน'])}") ; font=QFont("Terminal"); font.setPixelSize(8)
                Beamtag.setFont(font) ; Beamtag.setDefaultTextColor(QColor('cyan')) ; Beamtag.setRotation(-90)
                Beamtag.setPos(xOrg+pos1[0]-20,10+yOrg-pos1[1]-(pos2[1]-pos1[1])/2)
                self.scene.addItem(Beamtag)
                
    def closeEvent(self, event) :
        reply = QMessageBox.question(self,"ปิดหน้าต่าง","คุณแน่ใจไหมว่าต้องการออกจากหน้าต่างนี้",QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes :
            self.window_closed.emit()
            event.accept()
        else:
            event.ignore()
        


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainpath = os.getcwd()
    mainpath = mainpath.replace('\\','/')
    with open(f'{mainpath}/AMOLED.qss') as style :
        app.setStyleSheet(style.read())    
    window = BeamDataWindow(filename = r'C:\ProjectVenv\.venv\ProjectHouse\Deck.xlsx')
    window.show()
    sys.exit(app.exec())
    

