import pyqtgraph as pg
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import QFont,QColor,QPixmap,QImage,QPainter
import numpy as np
import sys
from indeterminatebeam import Beam, Support, PointLoadV, DistributedLoadV
import pandas as pd
import time
def draw_support_triangle(pos,size) :
    pos = pos
    size = size
    x = [-size+pos, pos, size+pos,-size+pos]
    y = [-size, 0, -size,-size]

    triangle = pg.PlotDataItem()
    triangle.setData(x=x, y=y, pen=pg.mkPen(color='pink',width = 1.5), fillLevel=0, fillBrush=(None))
    return triangle
def draw_support_line(pos,pos_size) :
    pos = pos
    pos_size = pos_size
    x = [pos-(pos_size*2),pos+(pos_size*2)]
    y = [-pos_size,-pos_size]
    line = pg.PlotDataItem()
    line.setData(x=x, y=y, pen=pg.mkPen(color='pink',width = 1.5), fillLevel=0, fillBrush=(None))
    return line     
def draw_beam(lenght,size) : #ใช้อันนี้ในการสร้าง คาน
        x = [0,lenght,lenght,0,0]
        y = [0,0,size,size,0]
        plot_beam = pg.PlotDataItem()
        plot_beam.setData(x=x, y=y, pen=pg.mkPen(color='cyan',width = 2.0,fillLevel = -1,fillBrush=(255, 0, 0, 100)))        
        return plot_beam
def draw_arrow(pos,lenght,color) :
    pos = pos
    lenght = lenght
    color = color
    x = [pos,pos]
    y = [0.1,lenght]
    arrow = pg.PlotDataItem()
    arrow.setData(x=x, y=y, pen=pg.mkPen(color=color,width = 1.5))        
    return arrow
def draw_arrow_head(pos,size,color) :
    pos = pos
    size = size
    color = color
    x = [pos-size,pos,pos+size]
    y = [size,0.1,size]
    arrow = pg.PlotDataItem()
    arrow.setData(x=x, y=y, pen=pg.mkPen(color=color,width = 1.5))        
    return arrow     
def draw_distribute_load(span,size) :
    start = span[0]
    end = span[1]
    x = [start,end,end,start,start]
    y = [0.1,0.1,size,size,0.1]
    distribute = pg.PlotDataItem()
    distribute.setData(x=x, y=y, pen=pg.mkPen(color='lightgreen',width = 1.5))        
    return distribute
def draw_dim_line(span) :
    start = span[0]
    end = span[1]
    x = [start,end]
    y = [-1,-1]

    dimline = pg.PlotDataItem()
    dimline.setData(x=x, y=y, pen=pg.mkPen(color='yellow',width = 1.5))        
    return dimline
def draw_vertical_dim_line(pos,lenght) :
    pos = pos
    lenght = lenght
    x = [pos,pos]
    y = [(-lenght/4)-1,(lenght/4)-1]

    ver_dimline = pg.PlotDataItem()
    ver_dimline.setData(x=x, y=y, pen=pg.mkPen(color='yellow',width = 1.5))        
    return ver_dimline
def dim_text(span) :
    offset = span[0]
    mid_pos = offset+(span[1]-span[0])/2
    dim_val = str((span[1]-span[0])*1000)

    text_item = pg.TextItem(text=dim_val, color='black', anchor=(0.5, 0.5))
    text_item.setPos(mid_pos, -0.8)
    return text_item
def val_text() :
    pass
def draw_support(obj,pos,node) : #ใช้อันนี้ในการสร้างซัพพอร์ต
    pos = pos
    node = str(node)
    obj.addItem(draw_support_triangle(pos,0.2))
    obj.addItem(draw_support_line(pos,0.2))

    text_item = pg.TextItem(text=node, color='pink', anchor=(0.5, 0.5))
    text_item.setPos(pos, -0.3)    
    obj.addItem(text_item)
def draw_point_load(obj,pos,node,force) : #ใช้อันนี้ในการสร้าง point load
    pos = pos
    node = f'{node:.0f}'
    force = f'{force:.3f} T'
    obj.addItem(draw_arrow(pos,0.6,'red'))
    obj.addItem(draw_arrow_head(pos,0.15,'red'))

    text_item = pg.TextItem(text=node, color='red', anchor=(0.5, 0.3))
    text_item.setPos(pos, -0.1)    
    obj.addItem(text_item)

    force_item = pg.TextItem(text=force, color='red', anchor=(0.5, 0.5))
    force_item.setPos(pos+0.5, 0.7)    
    obj.addItem(force_item)
def draw_dist_load(obj,span,force,size) : #ใช้อันนี้ในการสร้าง dist load
    color = 'lightgreen'
    size = size #หัวลูกศร default 0.2
    start = float(span[0])
    end = float(span[1])
    force = -force
    force = f'{force:.3f} T/M'
    mid_span = start+( (end-start) /2) #ตำแหน่งข้อความ
    obj.addItem(draw_distribute_load([start,end],size))
    obj.addItem(draw_arrow(start,size,color))
    obj.addItem(draw_arrow_head(start,0.12,color))
    obj.addItem(draw_arrow(end,size,color))
    obj.addItem(draw_arrow_head(end,0.12,color))    
    text_item = pg.TextItem(text=force, color=color, anchor=(0.5, 0.5))
    text_item.setPos(mid_span, size+0.1)    
    obj.addItem(text_item)
def draw_support_dim(obj,span) :
    start = span[0]
    end = span[1]
    lenght = round(span[1]-span[0],3)
    lenght = f'{lenght:.3f} m'
    mid_span = start+( (end-start) /2) #ตำแหน่งข้อความ
    obj.addItem(draw_dim_line([start,end]))
    obj.addItem(draw_vertical_dim_line(start,0.5))
    obj.addItem(draw_vertical_dim_line(end,0.5))
    text_item = pg.TextItem(text=lenght, color='yellow', anchor=(0.5, 0.5))
    text_item.setPos(mid_span, -0.9)    
    obj.addItem(text_item)
def plot_bm(obj,beam,Lenght,query_point) :
    text_item = pg.TextItem(text='B.M.D (T.M)', color='white', anchor=(0.5, 0.5))
    text_item.setPos(0, -2)    
    x = [0,0]
    y = [-2.5,-3.5]
    obj.plot(x=x, y=y, pen=pg.mkPen(color='red',width = 1.5))

    x = [Lenght,Lenght]
    y = [-2.5,-3.5]
    obj.plot(x=x, y=y, pen=pg.mkPen(color='red',width = 1.5))

    obj.addItem(text_item)
    x = beam._plotting_vectors['x']
    y = beam._plotting_vectors['bm']['y_vec']
    zipped_xy = list(zip(x,y))
    maxY = max(y,key=abs)
    y = [i/maxY for i in y]
    y = [i-3 for i in y]
    obj.plot(x, y, fillLevel=-3, fillBrush='lightblue', pen=pg.mkPen(color='yellow',width = 1.5)) 
    x = [0,Lenght]
    y = [-3,-3]
    obj.plot(x=x, y=y, pen=pg.mkPen(color='darkblue',width = 1.0))

    query_check = []
    for i in zipped_xy :
        
        if round(float(i[0]),3) in query_point and round(float(i[0]),3) not in query_check :
            val_y = float(i[1])
            print(val_y)
            text_val = pg.TextItem(text=f'{val_y:.3f}', color='yellow', anchor=(0.5, 0.5))
            if -val_y >= 0 :
                text_val.setPos(i[0], ((val_y)/maxY)-2.8)
            else : text_val.setPos(i[0], ((val_y)/maxY)-3.3)
            obj.addItem(text_val)
            query_check.append(round(float(i[0]),3))

def plot_sf(obj,beam,Lenght,query_point) : 
    text_item = pg.TextItem(text='S.F.D (T)   ', color='white', anchor=(0.5, 0.5))
    text_item.setPos(0, -4)

    x = [0,0]
    y = [-4.5,-5.5]
    obj.plot(x=x, y=y, pen=pg.mkPen(color='red',width = 1.5))

    x = [Lenght,Lenght]
    y = [-4.5,-5.5]
    obj.plot(x=x, y=y, pen=pg.mkPen(color='red',width = 1.5))

    obj.addItem(text_item) 

    x = beam._plotting_vectors['x']
    y = beam._plotting_vectors['sf']['y_vec'] 
    zipped_xy = list(zip(x,y))
    maxY = max(y,key=abs)
    y = [i/maxY for i in y]
    y = [i-5 for i in y]
    obj.plot(x, y, fillLevel=-5, fillBrush='lightgreen', pen=pg.mkPen(color='yellow',width = 1.5))

    x = [0,Lenght]
    y = [-5,-5]
    obj.plot(x=x, y=y, pen=pg.mkPen(color='yellow',width = 1.0))

    query_check = []
    for index , i in enumerate(zipped_xy) :
        if round(float(i[0]),3) in query_point :
            if round(float(i[0]),3) != 0 and round(float(i[0]),3) != Lenght and round(float(i[0]),3) not in query_check :
                val_yleft = zipped_xy[index-1][1]
                val_yright = zipped_xy[index+1][1]
                text_val1 = pg.TextItem(text=f'{val_yleft:.3f}', color='yellow', anchor=(0.5, 0.5))
                text_val1.setPos(i[0], ((val_yleft)/maxY)-4.8)

                text_val2 = pg.TextItem(text=f'{val_yright:.3f}', color='yellow', anchor=(0.5, 0.5))
                text_val2.setPos(i[0], ((val_yright)/maxY)-5.2)

                obj.addItem(text_val1)
                obj.addItem(text_val2)
                query_check.append(round(float(i[0]),3))
            elif i[0] == 0 or i[0] == Lenght :
                if i[0] == 0 :
                    val_yleft = zipped_xy[index+1][1]
                    text_val1 = pg.TextItem(text=f'{val_yleft:.3f}', color='yellow', anchor=(0.5, 0.5))
                    text_val1.setPos(i[0], ((val_yleft)/maxY)-4.8)
                    obj.addItem(text_val1)

                else :
                    val_yright = zipped_xy[index-1][1]
                    text_val2 = pg.TextItem(text=f'{val_yright:.3f}', color='yellow', anchor=(0.5, 0.5))
                    text_val2.setPos(i[0], ((val_yright)/maxY)-5.2)
                    obj.addItem(text_val2)

                query_check.append(round(float(i[0]),3))

def plot_diagram(obj,NO,filepath) :
    df = pd.read_excel(filepath,sheet_name=f'MB{NO}')
    obj.setTitle(f'วิเคราะห์คานหมายเลข {NO} ') 
    Lenght = float(df.loc[0,'Lenght'])
    obj.setYRange(-6,1)
    obj.setXRange(-0.5,(Lenght//1)+2)
    supNode = df['supNode'].dropna().tolist()
    supPos = df['supPos'].dropna().tolist()
    PLoad = df['PLoad'].dropna().tolist()
    PLoadNode = df['PLoadNode'].dropna().tolist()
    PLoadPos = df['PLoadPos'].dropna().tolist()
    LLoad = df['LLoad'].dropna().tolist()
    x0 = df['x0'].dropna().tolist()
    x1 = df['x1'].dropna().tolist()
    query_point = []
    [query_point.append(i) for i in supPos]
    [query_point.append(i) for i in PLoadPos]
    print('querypoint is',query_point)

    if LLoad != [] :
        sortLoad = sorted(LLoad,reverse=True)
    LLoadsize = []
    for i in LLoad :
        for j in sortLoad :
            if j == i :
                LLoadsize.append(sortLoad.index(j)+1)
                break
    print(LLoadsize)

    #สร้างคาน
    obj.addItem(draw_beam(Lenght,0.05))
    beam = Beam(float(Lenght))
    beam.update_units('force', 'T') ; beam.update_units('distributed', 'T/m') ; beam.update_units('moment', 'T.m') ; beam.update_units('E', 'kg/cm2') ; beam.update_units('I', 'mm4') ; beam.update_units('deflection', 'mm')

    #สร้างซัพพอร์ต
    for i in range(len(supNode)) :
        draw_support(obj,pos=supPos[i],node=f'{supNode[i]:.0f}')
        sup = Support(float(supPos[i]),(1,1,0),nodeid = f'{supNode[i]:.0f}')
        beam.add_supports(sup)
        if i != len(supNode) - 1 :
                draw_support_dim(obj,[float(supPos[i]),float(supPos[i+1])])
    if supPos[len(supPos)-1] and supPos[len(supPos)-1] != Lenght :
        draw_support_dim(obj,[float(supPos[len(supPos)-1] ),float(Lenght)])
            

    #สร้าง distLoad
    for i in range(len(LLoad)) :
        draw_dist_load(obj,span=[float(x0[i]),float(x1[i])],force = LLoad[i] ,size = 0.17+(LLoadsize[i]/25) )
        distload = DistributedLoadV(float(LLoad[i]),(float(x0[i]),float(x1[i])))
        beam.add_loads(distload)      

    if PLoad != [] :
        for i in range(len(PLoad)) :
            draw_point_load(obj,pos=PLoadPos[i],node=PLoadNode[i],force=-PLoad[i])
            pointload = PointLoadV(PLoad[i],PLoadPos[i])
            beam.add_loads(pointload)             

    beam.analyse()
    plot_bm(obj,beam,Lenght=Lenght,query_point=query_point)
    plot_sf(obj,beam,Lenght=Lenght,query_point=query_point)

   
class PlotWindow(QWidget) :
    window_closed = pyqtSignal()
    def __init__(self,filepath,DL,LL) :
        super().__init__()

        self.setWindowTitle("Auto Rc")
        self.setFixedSize(1300,700) #595 842
        self.setStyleSheet('background-color : None')
        self.central_widget = QWidget(self)
        self.layout = QVBoxLayout(self.central_widget)
        self.central_widget.move(25,90)
        self.plot_widget = pg.GraphicsLayoutWidget()  # This is the widget that will hold the plot
        self.plot_widget.setFixedSize(470,560)
        self.plot_widget.setBackground('white') #181818


        self.layout.addWidget(self.plot_widget)
        #plot_diagram
        self.plot_area = self.plot_widget.addPlot(title=f"วิเคราะห์คานหมายเลข {1}")
        self.plot_area.getAxis('bottom').setVisible(False)  # ซ่อนแกน X
        self.plot_area.getAxis('left').setVisible(False) 
        self.plot_area.setTitle('') #เปลี่ยนชื่อ title
        self.filepath = filepath
        plot_diagram(self.plot_area,1,self.filepath)
        self.Control_df = pd.read_excel(filepath.replace('-Analysed.xlsx','.xlsx'),sheet_name='Control_Parameter')
        self.BeamNo = int(self.Control_df.loc[0,'จำนวนคาน'])


        self.table_layout = QVBoxLayout()
        self.tableWidget = QTableWidget(self)

        self.tableWidget.setFixedSize(750,560) ; self.tableWidget.move(525,100)
        self.plot_area.getAxis('bottom').setTicks([])  ; self.plot_area.getAxis('left').setTicks([])
        self.tableWidget.setColumnWidth(0, 140)
        self.tableWidget.verticalHeader().setVisible(False)

        self.tableWidget.setStyleSheet("QHeaderView::section { background-color: purple; color: yellow; }")
        for i in range(0,8) :
            self.tableWidget.setColumnWidth(i, 50) 

        self.load_excel_data(1)
        self.table_layout.addWidget(self.tableWidget)
        

        self.nobeam = QComboBox(self) ; self.nobeam.addItems(str(i) for i in range(1,self.BeamNo+1)) ; self.nobeam.setFixedSize(140,30);self.nobeam.move(150,65)
        self.nobeam.currentIndexChanged.connect(self.UpdateValue)       

        lb = QLabel('คานที่แสดง :',self) ; font = QFont() ; font.setPointSize(10),lb.setFont(font) ;lb.setStyleSheet('color : yellow');  lb.move(60,70)
        title_lb = QLabel(f'ผลการวิเคราะห์และออกแบบคาน Load Combination {DL}DL + {LL}LL',self) ; titlefont = QFont() ; titlefont.setPointSize(20),title_lb.setFont(titlefont) ;title_lb.setStyleSheet('color : white');  title_lb.move(50,15)
        self.next_button = QPushButton("ถัดไป",self) ;self.next_button.setFont(font);self.next_button.setFixedSize(100,30) ;self.next_button.move(400,65)
        self.next_button.clicked.connect(self.UpdateValue)
        self.previous_button = QPushButton("ก่อนหน้า",self) ;self.previous_button.setFont(font);self.previous_button.setFixedSize(100,30) ;self.previous_button.move(300,65)
        self.previous_button.clicked.connect(self.UpdateValue)    


    def UpdateValue(self) :
        sender = self.sender() 
        if sender == self.next_button :
            nobeam = self.nobeam.currentIndex() + 1
            if nobeam < 14 :
                self.nobeam.setCurrentIndex(nobeam)
            else : nobeam - 1 

        elif sender == self.previous_button :
            nobeam = self.nobeam.currentIndex()-1
            if nobeam >= 0 :
                self.nobeam.setCurrentIndex(nobeam)
            else :
                nobeam = nobeam+1           

        nobeam = self.nobeam.currentIndex() + 1 


        image = QImage(self.plot_widget.size(), QImage.Format.Format_RGB888)
        new_width, new_height = image.width() * 10, image.height() * 10
        high_res_image = QImage(new_width, new_height, QImage.Format.Format_ARGB32)
        high_res_image.fill(0) 
        painter = QPainter(image)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_SourceOver)
        painter.drawImage(0, 0, image.scaled(new_width, new_height, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        self.plot_widget.render(painter)
        painter.end()
        image.save(f'C:\ProjectVenv\.venv\ProjectHouse\imageoutput\{nobeam}.png')

        self.plot_area.clear()

        plot_diagram(self.plot_area,nobeam,self.filepath)
        self.load_excel_data(nobeam)

    def load_excel_data(self,NO):
        df = pd.read_excel(self.filepath,sheet_name=f'Beam{NO}')
        df = df[['Span No','Section List','Moment','Top-As','Bot-As','Shear','RB6','RB9','Remark']]

        self.tableWidget.setRowCount(df.shape[0])
        self.tableWidget.setColumnCount(df.shape[1])
        self.tableWidget.setColumnWidth(0,175)
        for i in range(1,8) : self.tableWidget.setColumnWidth(i,67)
        self.tableWidget.setHorizontalHeaderLabels(df[['Span No','Section List','Moment','Top-As','Bot-As','Shear','RB6','RB9','Remark']])

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
    window = PlotWindow(filepath = r'C:\ProjectVenv\.venv\ProjectHouse\Deck-Analysed.xlsx',DL= 1.4,LL=1.7)
    window.show()

    sys.exit(app.exec())
from PyQt6.QtGui import QImage, QPainter, QImageReader
from PyQt6.QtCore import Qt
import numpy as np

