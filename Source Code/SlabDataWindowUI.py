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




class SlabDataWindow(QWidget) :
    window_closed = pyqtSignal()
    def __init__(self,filename) :
        super().__init__()
        self.setWindowTitle("Slab Data")
        self.setFixedSize(1300,600) #675 600
        self.setStyleSheet('background-color : None')
        self.view = QGraphicsView(self)
        self.view.setStyleSheet('background-color : wireframe')
        self.view.setFixedSize(QSize(600,550))
        self.view.setDragMode(self.view.DragMode.ScrollHandDrag)
        self.view.move(675,20)
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
            self.Slab_Data_df = pd.read_excel(self.filepath,sheet_name='Slab_Data')
        except :
            self.Slab_Data_df = pd.DataFrame({'หมายเลขพื้น' : [],'จุดต่อ I' : [],'จุดต่อ J' : [], 'จุดต่อ K' : [], 'จุดต่อ L' : [],'น้ำหนักบรรทุกคงที่อื่นๆ' : [],
                             'น้ำหนักบรรทุกจรอื่นๆ' : [],'ประเภทแผ่นพื้น':[],'ConnectionType' : [],'Ix' : [],'Iy' :[],'Jx' : [],'Jy':[],'Kx' : [],'Ky' : [],'Lx' : [],'Ly' : []})

        for i in range(int(self.Control_Parameter_df.loc[0,'จำนวนแผ่นพื้น'])) :
            self.Slab_Data_df.loc[i,'หมายเลขพื้น'] = str(i+1)

        OutFrame = QFrame(self) ; OutFrame.setFrameShape(QFrame.Shape.Box); OutFrame.setFrameShadow(QFrame.Shadow.Sunken) ; OutFrame.setLineWidth(3) 
        OutFrame.setFixedSize(630,560);OutFrame.setStyleSheet('border:5px solid purple') ; OutFrame.move(25,15)
        frame = QFrame(self) ; frame.setFrameShape(QFrame.Shape.Box); frame.setFrameShadow(QFrame.Shadow.Sunken) ; frame.setLineWidth(3) ;frame.setFixedSize(580,340);frame.setStyleSheet('background:rgba(128, 0, 128,120)') ; frame.move(50,160)
        for i, name in enumerate(['Project Title :','Floor Layer :','Engineer :','Date :']) :
            lb = QLabel(name,self) ; lb.setStyleSheet('color : purple') ; font = QFont() ; font.setPointSize(15) ; lb.setFont(font) ; lb.move(70,30*(i+1))

        #สร้างLabel จาก dataframe Head_And_Title
        for Column in range(self.Head_And_Title_df1.shape[1]) :
            lb = QLabel(f'{self.Head_And_Title_df1.iloc[0][Column]}',self) ; lb.setStyleSheet('color : green'); font = QFont() ; font.setPointSize(15) ; lb.setFont(font) ; lb.move(210,30*(Column+1))
        
        for i , name in enumerate(['จำนวนแผ่น','แผ่นที่...']) :
            lb = QLabel(name,self) ; lb.setStyleSheet('color : white') ; font = QFont() ; font.setPointSize(14) ; lb.setFont(font) ; lb.move(130,190+(35*(i+1)))
            if name == 'จำนวนแผ่น' :
                noslabData =  str(self.Control_Parameter_df.loc[0,'จำนวนแผ่นพื้น'])
                noslabDataEntry = QLineEdit(noslabData,self);noslabDataEntry.setReadOnly(True) ;noslabDataEntry.setStyleSheet('color : yellow'); noslabDataEntry.setAlignment(Qt.AlignmentFlag.AlignHCenter);noslabDataEntry.setFixedSize(100,30);noslabDataEntry.move(230,190+(35*(i+1)))
                pass
            elif name == 'แผ่นที่...' :
                self.slabCombo = QComboBox(self) ; self.slabCombo.addItems(str(i) for i in range(1,int(self.Control_Parameter_df.loc[0,'จำนวนแผ่นพื้น'])+1));self.slabCombo.setFixedSize(100,30);self.slabCombo.move(230,190+(35*(i+1)))

                pass
        
        for i , name in enumerate(['จุดต่อ I','จุดต่อ J','จุดต่อ K','จุดต่อ L']) :
            lb = QLabel(name,self) ; lb.setStyleSheet('color : yellow') ; font = QFont() ; font.setPointSize(12) ; lb.setFont(font) ; lb.move(140,330+(30*(i+1)))
            globals()[f'NodeSlab{i}'] = QLineEdit(self); globals()[f'NodeSlab{i}'].setText(f'{str(self.Slab_Data_df.iloc[self.slabCombo.currentIndex()][i+1])}') ; globals()[f'NodeSlab{i}'] ; globals()[f'NodeSlab{i}'].setAlignment(Qt.AlignmentFlag.AlignHCenter);globals()[f'NodeSlab{i}'].setFixedSize(80,20);globals()[f'NodeSlab{i}'].move(220,333+(30*(i+1)))
            
        for i , name in enumerate(['Other Dead Load(T/M)','Other Live Load(T/M)','ประเภทแผ่นพื้น']) :
            lb = QLabel(name,self) ; lb.setStyleSheet('color : yellow') ; font = QFont() ; font.setPointSize(12) ; lb.setFont(font) ; lb.move(350,330+(30*(i+1)))
            if name == 'Other Dead Load(T/M)' :
                self.OtherDeadLoad = QLineEdit(str(self.Slab_Data_df.loc[i,'น้ำหนักบรรทุกคงที่อื่นๆ']),self) ; self.OtherDeadLoad ; self.OtherDeadLoad.setAlignment(Qt.AlignmentFlag.AlignHCenter);self.OtherDeadLoad.setFixedSize(80,20);self.OtherDeadLoad.move(510,330+(30*(i+1)))
            elif name == 'Other Live Load(T/M)' :
                self.OtherLiveLoad = QLineEdit(str(self.Slab_Data_df.loc[i,'น้ำหนักบรรทุกจรอื่นๆ']),self) ; self.OtherLiveLoad ; self.OtherLiveLoad.setAlignment(Qt.AlignmentFlag.AlignHCenter);self.OtherLiveLoad.setFixedSize(80,20);self.OtherLiveLoad.move(510,330+(30*(i+1)))
            elif name == 'ประเภทแผ่นพื้น' :
                self.ListData = ['แผ่นพื้นสำเร็จวางแนวตั้ง','แผ่นพื้นสำเร็จวางแนวนอน','พื้นหล่อในที่']
                self.TypeSlab = QComboBox(self) ; self.TypeSlab.addItems(i for i in self.ListData) ; self.TypeSlab.setFixedSize(160,30);self.TypeSlab.move(460,330+(30*(i+1)))

        #นำเข้ารูปภาพประกอบแผ่นพื้น
        self.TypeSlab.setCurrentText(str(self.Slab_Data_df.loc[self.slabCombo.currentIndex(),'ประเภทแผ่นพื้น']))
        imgpath = os.getcwd()
        imgpath = imgpath.replace('\\','/')
        img = QPixmap(f'{imgpath}/slabimg.jpg')
        label = QLabel(self)
        label.setPixmap(img)
        label.move(380,180)
        #save button
        self.savebtn = QPushButton('บันทึก',self) ; self.savebtn.setFixedSize(QSize(160,50)) ; self.savebtn.move(450,510)
        #signal
        self.slabCombo.currentIndexChanged.connect(self.UpdateValue)
        self.savebtn.clicked.connect(self.SaveData)


    def UpdateValue(self) :
        print(self.slabCombo.currentIndex())
        numslab = self.slabCombo.currentIndex()
        globals()[f'NodeSlab{0}'].setText(str(self.Slab_Data_df.loc[numslab,'จุดต่อ I']))
        globals()[f'NodeSlab{1}'].setText(str(self.Slab_Data_df.loc[numslab,'จุดต่อ J']))
        globals()[f'NodeSlab{2}'].setText(str(self.Slab_Data_df.loc[numslab,'จุดต่อ K']))
        globals()[f'NodeSlab{3}'].setText(str(self.Slab_Data_df.loc[numslab,'จุดต่อ L']))
        self.OtherLiveLoad.setText(str(self.Slab_Data_df.loc[numslab,'น้ำหนักบรรทุกจรอื่นๆ']))
        self.OtherDeadLoad.setText(str(self.Slab_Data_df.loc[numslab,'น้ำหนักบรรทุกคงที่อื่นๆ']))
        self.TypeSlab.setCurrentText(str(self.Slab_Data_df.loc[numslab,'ประเภทแผ่นพื้น']))
    def SaveData(self) :
        self.view.centerOn(200, 750)
        self.scene.clear()
        IndexNumslab = self.slabCombo.currentIndex()
        indexTypeSlab = self.TypeSlab.currentIndex()
        self.Slab_Data_df.loc[IndexNumslab,'จุดต่อ I'] = int(globals()[f'NodeSlab{0}'].text())
        self.Slab_Data_df.loc[IndexNumslab,'จุดต่อ J'] = int(globals()[f'NodeSlab{1}'].text())
        self.Slab_Data_df.loc[IndexNumslab,'จุดต่อ K'] = int(globals()[f'NodeSlab{2}'].text())
        self.Slab_Data_df.loc[IndexNumslab,'จุดต่อ L'] = int(globals()[f'NodeSlab{3}'].text())
        self.Slab_Data_df.loc[IndexNumslab,'น้ำหนักบรรทุกคงที่อื่นๆ'] = self.OtherDeadLoad.text()
        self.Slab_Data_df.loc[IndexNumslab,'น้ำหนักบรรทุกจรอื่นๆ'] = self.OtherLiveLoad.text()
        self.Slab_Data_df.loc[IndexNumslab,'ประเภทแผ่นพื้น'] = self.ListData[indexTypeSlab]

        NodeI = int(self.Slab_Data_df.loc[IndexNumslab,'จุดต่อ I'])
        NodeJ = int(self.Slab_Data_df.loc[IndexNumslab,'จุดต่อ J'])
        NodeK = int(self.Slab_Data_df.loc[IndexNumslab,'จุดต่อ K'])
        NodeL = int(self.Slab_Data_df.loc[IndexNumslab,'จุดต่อ L'])
        for index,Node_i in enumerate(self.Node_df['Node No. (int)']) :
            if Node_i == NodeI :
                self.Slab_Data_df.loc[IndexNumslab,'Ix'] = self.Node_df.loc[index,'พิกัด X (m)(.2float)']
                self.Slab_Data_df.loc[IndexNumslab,'Iy'] = self.Node_df.loc[index,'พิกัด Y (m)  (.2float)']
            elif Node_i == NodeJ :
                self.Slab_Data_df.loc[IndexNumslab,'Jx'] = self.Node_df.loc[index,'พิกัด X (m)(.2float)']
                self.Slab_Data_df.loc[IndexNumslab,'Jy'] = self.Node_df.loc[index,'พิกัด Y (m)  (.2float)']
            elif Node_i == NodeK :
                self.Slab_Data_df.loc[IndexNumslab,'Kx'] = self.Node_df.loc[index,'พิกัด X (m)(.2float)']
                self.Slab_Data_df.loc[IndexNumslab,'Ky'] = self.Node_df.loc[index,'พิกัด Y (m)  (.2float)']
            elif Node_i == NodeL :
                self.Slab_Data_df.loc[IndexNumslab,'Lx'] = self.Node_df.loc[index,'พิกัด X (m)(.2float)']
                self.Slab_Data_df.loc[IndexNumslab,'Ly'] = self.Node_df.loc[index,'พิกัด Y (m)  (.2float)']

        
        self.SlabDraw()
        self.PltNode()
        self.DimDraw()
        save_excel_sheet(self.Slab_Data_df, self.filepath, sheetname='Slab_Data', index=False)
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
        elif event.key() == Qt.Key.Key_X:
            self.view.scale(1/1.1,1/1.1)
        elif event.key() == Qt.Key.Key_Z:
            self.view.scale(1.1,1.1)
        super().keyPressEvent(event)

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
                Dimtext.setDefaultTextColor(QColor('gold')) ; Dimtext.setPos(210-offset,(20+700-X+700-X+(dx2x1))/2)
                self.scene.addItem(Hdim)
                self.scene.addItem(Dimtext)
                vline1 = QGraphicsLineItem(225-offset, 700-X,235-offset, 700-X) ; vline1.setPen(QPen(Qt.GlobalColor.darkGreen))
                vline2 = QGraphicsLineItem(225-offset, 700-X+dx2x1 ,235-offset, 700-X+dx2x1 ) ; vline2.setPen(QPen(Qt.GlobalColor.darkGreen))
                self.scene.addItem(vline1) ;self.scene.addItem(vline2)


        for i , val in enumerate(y) :
            if i != 0 :
                offset = -min(y) * 20
                X = round(val,3)*20
                dy2y1 = round(y[i]-y[i-1],3) *20
                Hdim = QGraphicsLineItem(300+X, 1000-230+offset,300+X-dy2y1,1000-230+offset) ; Hdim.setPen(QPen(Qt.GlobalColor.darkGreen))
                Dimtext = QGraphicsTextItem(str(round((dy2y1/20),3))) ; font.setPixelSize(8) ; Dimtext.setFont(font)
                Dimtext.setDefaultTextColor(QColor('gold')) ; Dimtext.setPos(-12+300+X-((dy2y1)/2),1000-225+offset)
                self.scene.addItem(Hdim)
                self.scene.addItem(Dimtext)
                vline1 = QGraphicsLineItem(300+X, 1000-225+offset , 300+X, 1000-235+offset ) ; vline1.setPen(QPen(Qt.GlobalColor.darkGreen))
                vline2 = QGraphicsLineItem(300+X-dy2y1, 1000-225+offset ,300+X-dy2y1, 1000-235+offset ) ; vline2.setPen(QPen(Qt.GlobalColor.darkGreen))
                self.scene.addItem(vline1) ;self.scene.addItem(vline2)



            
    def SlabDraw(self) :
        for i,value in enumerate(self.Slab_Data_df['หมายเลขพื้น']) :
            Ix = float(self.Slab_Data_df.loc[i,'Ix'])*20
            Iy = float(self.Slab_Data_df.loc[i,'Iy'])*20
            Jx = float(self.Slab_Data_df.loc[i,'Jx'])*20
            Jy = float(self.Slab_Data_df.loc[i,'Jy'])*20
            Kx = float(self.Slab_Data_df.loc[i,'Kx'])*20
            Ky = float(self.Slab_Data_df.loc[i,'Ky'])*20
            Lx = float(self.Slab_Data_df.loc[i,'Lx'])*20
            Ly = float(self.Slab_Data_df.loc[i,'Ly'])*20
            tagposX = (Ix+Jx)/2
            tagposY = (Ly+Iy)/2
            recttest = QGraphicsRectItem(300+Lx, 700-Ly, Jx-Ix, Ly-Iy) ; recttest.setPen(QPen(QColor(0,255,0)))
            self.scene.addItem(recttest)

            if self.Slab_Data_df.loc[i,'ประเภทแผ่นพื้น'] == 'พื้นหล่อในที่' :
                line_item1 = QGraphicsLineItem(300+Ix, 700-Iy, 300+Kx, 700-Ky) ; line_item1.setPen(QPen(QColor(0,255,0)))
                line_item2 = QGraphicsLineItem(300+Lx, 700-Ly, 300+Jx, 700-Jy) ; line_item2.setPen(QPen(QColor(0,255,0)))
                self.scene.addItem(line_item1)
                self.scene.addItem(line_item2)
                slabtag = QGraphicsTextItem(str('S')+str(self.Slab_Data_df.loc[i,'หมายเลขพื้น'])) ; font=QFont("Terminal"); font.setPixelSize(8)

                Circle_tag = QGraphicsEllipseItem(300+tagposX-6.25,700-tagposY-6.25,12.5,12.5) ; Circle_tag.setBrush(QBrush(QColor(0,0,0)))
                Circle_tag.setPen(QPen(Qt.GlobalColor.black))
                slabtag.setPos(300+tagposX-10,700-tagposY-7.5)
                slabtag.setFont(font) ; slabtag.setDefaultTextColor(QColor('lightgreen'))
                self.scene.addItem(Circle_tag)
                self.scene.addItem(slabtag)
                pass
            elif self.Slab_Data_df.loc[i,'ประเภทแผ่นพื้น'] == 'แผ่นพื้นสำเร็จวางแนวตั้ง' :
                line_item1 = QGraphicsLineItem(300+tagposX, 700-tagposY+15, 300+tagposX, 700-tagposY-15) ; line_item1.setPen(QPen(QColor(0,255,0)))
                line_itemarr1 = QGraphicsLineItem(300+tagposX, 700-tagposY+15, 300+tagposX+3, 700-tagposY+10) ; line_itemarr1.setPen(QPen(QColor(0,255,0)))
                line_itemarr2 = QGraphicsLineItem(300+tagposX, 700-tagposY-15, 300+tagposX-3, 700-tagposY-10) ; line_itemarr2.setPen(QPen(QColor(0,255,0)))
                Circle_tag = QGraphicsEllipseItem(300+tagposX-6.25,700-tagposY-6.25,12.5,12.5) ; Circle_tag.setBrush(QBrush(QColor(0,0,0)))
                slabtag = QGraphicsTextItem(str('S')+str(self.Slab_Data_df.loc[i,'หมายเลขพื้น'])) ; font=QFont("Terminal"); font.setPixelSize(8)
                slabtag.setFont(font) ; slabtag.setDefaultTextColor(QColor('lightgreen'))
                Circle_tag.setPen(QPen(Qt.GlobalColor.black))
                slabtag.setPos(300+tagposX-10,700-tagposY-7.5)
                self.scene.addItem(line_item1)
                self.scene.addItem(line_itemarr1)
                self.scene.addItem(line_itemarr2)
                self.scene.addItem(Circle_tag)
                self.scene.addItem(slabtag)
            elif self.Slab_Data_df.loc[i,'ประเภทแผ่นพื้น'] == 'แผ่นพื้นสำเร็จวางแนวนอน' :
                line_item1 = QGraphicsLineItem(300+tagposX-15, 700-tagposY, 300+tagposX+15, 700-tagposY) ; line_item1.setPen(QPen(QColor(0,255,0)))
                line_itemarr1 = QGraphicsLineItem(300+tagposX+15, 700-tagposY, 300+tagposX+10, 700-tagposY-3) ; line_itemarr1.setPen(QPen(QColor(0,255,0)))
                line_itemarr2 = QGraphicsLineItem(300+tagposX-15, 700-tagposY, 300+tagposX-10, 700-tagposY+3) ; line_itemarr2.setPen(QPen(QColor(0,255,0)))
                Circle_tag = QGraphicsEllipseItem(300+tagposX-6.25,700-tagposY-6.25,12.5,12.5) ; Circle_tag.setBrush(QBrush(QColor(0,0,0)))
                slabtag = QGraphicsTextItem(str('S')+str(self.Slab_Data_df.loc[i,'หมายเลขพื้น'])) ; font=QFont("Terminal"); font.setPixelSize(8)
                slabtag.setFont(font) ; slabtag.setDefaultTextColor(QColor('lightgreen'))
                Circle_tag.setPen(QPen(Qt.GlobalColor.black))
                slabtag.setPos(300+tagposX-10,700-tagposY-7.5)
                self.scene.addItem(line_item1)
                self.scene.addItem(line_itemarr1)
                self.scene.addItem(line_itemarr2)
                self.scene.addItem(Circle_tag)
                self.scene.addItem(slabtag)



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
            self.scene.addItem(recttest)
 
    def closeEvent(self, event) :
        reply = QMessageBox.question(self,"ปิดหน้าต่าง","คุณแน่ใจไหมว่าต้องการออกจากหน้าต่างนี้",QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes :
            self.window_closed.emit()
            event.accept()
        else:
            event.ignore()          


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SlabDataWindow(filename = r'C:\ProjectVenv\.venv\ProjectHouse\Deck.xlsx')
    window.show()
    sys.exit(app.exec())