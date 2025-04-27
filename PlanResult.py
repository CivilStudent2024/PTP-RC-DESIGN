from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.colors import red, blue, green,black
from reportlab.lib import colors
import pandas as pd
import os 

import math

def PltNode(obj,Node_df,offset_x,offset_y,scale) :
    Xorg = (841.89/2) - ((offset_x)*(28.35/scale)) ; Yorg =  ((offset_y)*(28.35/scale))
    for i,value in enumerate(Node_df['Node No. (int)']) :
        obj.setFont("Helvetica", 10)
        obj.setLineWidth(0.35)
        x = float(Node_df.loc[i,'พิกัด X (m)(.2float)']) * (28.35/scale) ;y = float(Node_df.loc[i,'พิกัด Y (m)  (.2float)']) *(28.35/scale)

        if Node_df.loc[i,'สถานะจุดต่อ'] == 'Y' :
            obj.rect(Xorg+x-(2.835/scale), Yorg+y-(2.835/scale), (5.67/scale), (5.67/scale),stroke=1,fill=1) 
        elif Node_df.loc[i,'สถานะจุดต่อ'] == 'N' :
            obj.rect(Xorg+x-(2.835/scale), Yorg+y-(2.835/scale), (5.67/scale), (5.67/scale),stroke=1,fill=0) 
        elif Node_df.loc[i,'สถานะจุดต่อ'] == 'E' :
            obj.rect(Xorg+x-(2.835/scale), Yorg+y-(2.835/scale), (5.67/scale), (5.67/scale),stroke=1,fill=0) 
            obj.line(Xorg+x-(2.835/scale), Yorg+y-(2.835/scale),Xorg+x+(2.835/scale),Yorg+y+(2.835/scale))
            obj.line(Xorg+x+(2.835/scale), Yorg+y-(2.835/scale),Xorg+x-(2.835/scale),Yorg+y+(2.835/scale))
        obj.setFillColor(blue)
        obj.drawString(Xorg+x, Yorg+y+5, str(Node_df.loc[i,'Node No. (int)']))
        obj.setFillColor(black)
        obj.setLineWidth(1)

def BeamDraw(obj,Beam_df,offset_x,offset_y,scale) :
    for i,value in enumerate(Beam_df['ทิศทางการวาง']) :
        Xorg = (841.89/2) - ((offset_x)*(28.35/scale))
        Yorg =  ((offset_y)*(28.35/scale))
        
        if value == 'X' :
            pos1 = [Beam_df.loc[i,'Node1x'] * (28.35/scale) ,Beam_df.loc[i,'Node1y'] * (28.35/scale)] 
            pos2 = [Beam_df.loc[i,'Node2x'] * (28.35/scale) ,Beam_df.loc[i,'Node2y'] * (28.35/scale)] 
            obj.setLineWidth(0.25)
            obj.rect(Xorg+pos1[0]-(2.835/scale/1.3), Yorg+pos1[1]-(2.835/scale/1.3), (pos2[0]-pos1[0]) , (5.67/scale/1.3),stroke=1,fill=0) 
            obj.setFont("Helvetica", 8) ; obj.setFillColor(green)
            obj.drawString(Xorg+pos1[0]+(pos2[0]-pos1[0])/2, 10+pos1[1]+Yorg,f"B{(Beam_df.loc[i,'หมายเลขคาน'])}") 
            obj.setFillColor(black)
            obj.setLineWidth(1)

        elif value == 'Y' :

            pos1 = [Beam_df.loc[i,'Node1x'] * (28.35/scale) ,Beam_df.loc[i,'Node1y'] * (28.35/scale)] 
            pos2 = [Beam_df.loc[i,'Node2x'] * (28.35/scale) ,Beam_df.loc[i,'Node2y'] * (28.35/scale)]
            obj.setLineWidth(0.25)
            obj.rect(Xorg+pos1[0]-(2.835/scale/1.3), Yorg+pos1[1]-(2.835/scale/1.3), (5.67/scale/1.3) , (pos2[1]-pos1[1]),stroke=1,fill=0)
            obj.setFont("Helvetica", 8) ; obj.setFillColor(green)
            obj.translate(Xorg+pos1[0] + pos1[1]+(abs(pos1[1]-pos2[1])/2)+Yorg, pos1[1]+(abs(pos1[1]-pos2[1])/2)+Yorg)
            obj.rotate(90)
            obj.drawString(0, pos1[1]+(abs(pos1[1]-pos2[1])/2)+Yorg+10,f"B{(Beam_df.loc[i,'หมายเลขคาน'])}")
            obj.setFillColor(black)
            obj.rotate(-90)
            obj.translate(-(Xorg+pos1[0] + pos1[1]+(abs(pos1[1]-pos2[1])/2)+Yorg), -(pos1[1]+(abs(pos1[1]-pos2[1])/2)+Yorg))
            obj.setLineWidth(1)
            
def DimDraw(obj,Node_df,offset_x,offset_y,scale) :
    Xorg = (841.89/2) - ((offset_x)*(28.35/scale))
    Yorg =  ((offset_y)*(28.35/scale))
    x = Node_df['พิกัด Y (m)  (.2float)'].tolist()
    x = list(set(x))
    x.sort()
    y = Node_df['พิกัด X (m)(.2float)'].tolist()
    y = list(set(y))
    y.sort()
    print(y)

    for i , val in enumerate(x) :
        if i != 0 :
            X = round(val,3)* (28.35/scale)
            dx2x1 = round(x[i]-x[i-1],3) * (28.35/scale)
            offset = (-min(x) * (28.35/scale))-(4*(28.35/scale))
            obj.line(Xorg+offset ,Yorg+X,Xorg+offset,Yorg+X-dx2x1)
            obj.line(Xorg+offset+(28.35/scale*0.5) ,Yorg+X,Xorg+offset-(28.35/scale*0.5),Yorg+X)
            obj.line(Xorg+offset+(28.35/scale*0.5) ,Yorg+X-dx2x1,Xorg+offset-(28.35/scale*0.5),Yorg+X-dx2x1)
            obj.translate(Xorg+offset-(28.35/scale*0.5)+Yorg+X-(dx2x1/2),Yorg+X-(dx2x1/2)-10)
            obj.rotate(90)
            obj.drawString(0, Yorg+X-(dx2x1/2),f"{dx2x1/(28.35/scale):.2f} M")
            obj.rotate(-90)
            obj.translate(-(Xorg+offset-(28.35/scale*0.5)+Yorg+X-(dx2x1/2)),-(Yorg+X-(dx2x1/2)-10))

    for i , val in enumerate(y) :
        if i != 0 :
            X = round(val,3)*(28.35/scale)
            dy2y1 = round(y[i]-y[i-1],3) * (28.35/scale)
            offset = (-min(y) * (28.35/scale))-(4*(28.35/scale))
            obj.line(Xorg+X ,Yorg+offset,Xorg+X-dy2y1,Yorg+offset)
            obj.line(Xorg+X ,Yorg+offset+(28.35/scale*0.5),Xorg+X,Yorg+offset-(28.35/scale*0.5))
            obj.line(Xorg+X-dy2y1 ,Yorg+offset+(28.35/scale*0.5),Xorg+X-dy2y1,Yorg+offset-(28.35/scale*0.5))
            obj.drawString(Xorg+X-(dy2y1/2)-10, Yorg+offset-15,f"{dy2y1/(28.35/scale):.2f} M")


def LinloadDraw(obj,LineLoad_df,offset_x,offset_y,scale) :
    for i,value in enumerate(LineLoad_df['direction']) :
        Xorg = (841.89/2) - ((offset_x)*(28.35/scale))
        Yorg = ((offset_y)*(28.35/scale))
        pos1 = [float(LineLoad_df.loc[i,'Node1x']) * (28.35/scale) ,float(LineLoad_df.loc[i,'Node1y']) * (28.35/scale) ] 
        pos2 = [float(LineLoad_df.loc[i,'Node2x']) * (28.35/scale) ,float(LineLoad_df.loc[i,'Node2y']) * (28.35/scale) ] 
        obj.setLineWidth(0.3)
        if value == 'X' :

            lenload = (pos2[0]-pos1[0])/3
            for j in range(math.floor(lenload)) :
                obj.setStrokeColor(colors.red)
                obj.line(Xorg+pos1[0]-(2.835/scale/1.1)+(3*j), Yorg+pos1[1]-(2.835/scale/1.1),Xorg+pos1[0]+(2.835/scale)+((3*j)+1), Yorg+pos1[1]+(2.835/scale/1.1) )
                
                obj.setStrokeColor(colors.black)

            obj.setFillColor(red)
            obj.drawString(-10+Xorg+pos1[0]+(pos2[0]-pos1[0])/2,+10+Yorg+pos1[1], f"{LineLoad_df.loc[i,'Load(T/M)']} T/M")
            obj.setFillColor(black)
        else :
            lenload = (pos2[1]-pos1[1])/3
            for j in range(math.floor(lenload)) :
                #loadline = QGraphicsLineItem(xOrg+pos1[0]+1, yOrg-pos1[1]+1-(3*j),xOrg+pos1[0]-1, yOrg-pos1[1]-1-((3*j)+1) ); loadline.setPen(QPen(Qt.GlobalColor.red))
                obj.setStrokeColor(colors.red)
                
                obj.line(Xorg+pos1[0]-(2.835/scale/1.1), Yorg+pos1[1]-(2.835/scale/1.1)+(3*j),Xorg+pos1[0]+(2.835/scale), Yorg+pos1[1]+(2.835/scale/1.1)+((3*j)+1) )
                obj.setStrokeColor(colors.black)
            obj.setFillColor(red)
            obj.translate(Xorg+pos1[0]+Yorg+pos1[1]+(abs(pos2[1]-pos1[1])/2),Yorg+pos1[1]+(abs(pos2[1]-pos1[1])/2)-(28.35/scale*0.5))
            obj.rotate(90)
            obj.drawString(0, Yorg+pos1[1]+(abs(pos2[1]-pos1[1])/2)+10,f"{LineLoad_df.loc[i,'Load(T/M)']} T/M")
            obj.rotate(-90)
            obj.translate(-(Xorg+pos1[0]+Yorg+pos1[1]+(abs(pos2[1]-pos1[1])/2)),-(Yorg+pos1[1]+(abs(pos2[1]-pos1[1])/2)-(28.35/scale*0.5)))
            obj.setFillColor(black)
        obj.setLineWidth(1)

def PointLoadDraw(obj,Reaction_df,Node_df,offset_x,offset_y,scale) :
    Xorg = (841.89/2) - ((offset_x)*(28.35/scale)) ; Yorg =  ((offset_y)*(28.35/scale))
    obj.setFont("Helvetica", 10)
    for i,value in enumerate(Reaction_df['Node']) :
        for index,node in enumerate(Node_df['Node No. (int)']) : 
            if node == value :
                x = float(Node_df.loc[index,'พิกัด X (m)(.2float)']) * (28.35/scale) ;y = float(Node_df.loc[index,'พิกัด Y (m)  (.2float)']) *(28.35/scale)
                break
        obj.setFillColor(red)
        try :            
            obj.drawString(Xorg+x, Yorg+y-10, f"{Reaction_df.loc[i,'Load(T/M)']:.2f}")
        except :
            obj.drawString(Xorg+x, Yorg+y-10, f" ")

        obj.setFillColor(black)

def SlabDraw(obj,Slab_Data_df,Slab_Result_df,offset_x,offset_y,scale) :
    Xorg = (841.89/2) - ((offset_x)*(28.35/scale)) ; Yorg =  ((offset_y)*(28.35/scale))
    obj.setFont("Helvetica", 10)
    obj.setLineWidth(0.25)
    for i,value in enumerate(Slab_Data_df['หมายเลขพื้น']) :
        Ix = float(Slab_Data_df.loc[i,'Ix'])*(28.35/scale)
        Iy = float(Slab_Data_df.loc[i,'Iy'])*(28.35/scale)
        Jx = float(Slab_Data_df.loc[i,'Jx'])*(28.35/scale)
        Jy = float(Slab_Data_df.loc[i,'Jy'])*(28.35/scale)
        Kx = float(Slab_Data_df.loc[i,'Kx'])*(28.35/scale)
        Ky = float(Slab_Data_df.loc[i,'Ky'])*(28.35/scale)
        Lx = float(Slab_Data_df.loc[i,'Lx'])*(28.35/scale)
        Ly = float(Slab_Data_df.loc[i,'Ly'])*(28.35/scale)
        tagposX = (Ix+Jx)/2
        tagposY = (Ly+Iy)/2
        obj.rect(Xorg+Ix, Yorg+Iy, Jx-Ix, Ly-Iy,stroke=1,fill=0) ;
        obj.drawString(-5+Xorg+tagposX,Yorg+tagposY+5, f"S{value}") 
        for j,row in Slab_Result_df.iterrows() :
            if row['หมายเลขพื้น'] == value :
                try :
                    obj.drawString(-15+Xorg+tagposX,Yorg+tagposY-5, f"({row['Thickness']:.3f})")
                except :
                    obj.drawString(-20+Xorg+tagposX,Yorg+tagposY-5, f"(plank slab)")
                break
    obj.setLineWidth(1)
def ScaleTextDraw(obj,scale,text=None) :
    text = str(text)
    obj.setFont("Helvetica", (28.35/scale*0.5))
    obj.drawString(700,75,f"{text}") 
    obj.drawString(700,50,f"Scale 1:{scale*100:.0f}") 
    obj.setFont("Helvetica", 10)

def Plan_Write(filepath,file_name) :
    Node_df = pd.read_excel(filepath,sheet_name='Node_Data')
    Beam_df = pd.read_excel(filepath,sheet_name='Beam_Data')
    LineLoad_df = pd.read_excel(filepath,sheet_name='Lineload_Data')
    Slab_Data_df = pd.read_excel(filepath,sheet_name='Slab_Data')
    Reaction_df = pd.read_excel(filepath.replace('.xlsx','-Analysed.xlsx'),sheet_name='Reaction_result')
    Slab_Result_df = pd.read_excel(filepath.replace('.xlsx','-Analysed.xlsx'),sheet_name='Slab_Result')
    Slab_Result_df.fillna('',inplace=True)
    # สร้าง PDF
    report_save_path = os.getcwd() + '\Report'


    c = canvas.Canvas(f'{report_save_path}\\{file_name}-Plan.pdf', pagesize=landscape(A4))

    scale = 1.00
    x = Node_df['พิกัด Y (m)  (.2float)'].tolist()
    x = list(set(x))
    x.sort()
    y = Node_df['พิกัด X (m)(.2float)'].tolist()
    y = list(set(y))
    y.sort()

    offset_x = (max(y)- min(y))/2
    offset_y = (max(x)- min(x))/2
    if offset_x * 2 >= 20 or offset_y *2 >= 12 :
        scale = 1.25
    if offset_x * 2 >= 20 or offset_y *2 >= 20 :
        scale = 1.50

    print(offset_x*2)
    print(offset_y*2)
    #แปลนคาน
    BeamDraw(obj = c,Beam_df= Beam_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    PltNode(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    DimDraw(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    ScaleTextDraw(obj = c,scale =scale,text='Beam Plan')


    #แปลน pointload
    c.showPage()
    PltNode(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    DimDraw(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    PointLoadDraw(obj = c,Reaction_df= Reaction_df,Node_df =  Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    ScaleTextDraw(obj = c,scale =scale,text='Node Plan')

    #แปลน distribute load
    c.showPage()
    BeamDraw(obj = c,Beam_df= Beam_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    PltNode(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    DimDraw(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    LinloadDraw(obj = c,LineLoad_df = LineLoad_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    ScaleTextDraw(obj = c,scale =scale,text= '')
    c.showPage()
    DimDraw(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    SlabDraw(obj = c,Slab_Data_df=Slab_Data_df,Slab_Result_df=Slab_Result_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    PltNode(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    ScaleTextDraw(obj = c,scale =scale,text= 'Slab Plan')
    # บันทึกและปิดไฟล์
    c.save()

    print(f"สร้างไฟล์ PDF: {file_name} แล้ว ")


if __name__ == '__main__' :   

    Plan_Write(filepath = r'C:\ProjectVenv\.venv\ProjectHouse\Arr1.xlsx' ,file_name = 'Arr1')
    


    """
    filepath = r'C:\ProjectVenv\.venv\ProjectHouse\Exhouse.xlsx'
    Node_df = pd.read_excel(filepath,sheet_name='Node_Data')
    Beam_df = pd.read_excel(filepath,sheet_name='Beam_Data')
    LineLoad_df = pd.read_excel(filepath,sheet_name='Lineload_Data')
    Slab_Data_df = pd.read_excel(filepath,sheet_name='Slab_Data')
    Reaction_df = pd.read_excel(filepath.replace('.xlsx','-Analysed.xlsx'),sheet_name='Reaction_result')
    Slab_Result_df = pd.read_excel(filepath.replace('.xlsx','-Analysed.xlsx'),sheet_name='Slab_Result')
    Slab_Result_df.fillna('',inplace=True)
    # สร้าง PDF
    pdf_filepath = "rectangle_example.pdf"
    c = canvas.Canvas(pdf_filepath, pagesize=landscape(A4))

    scale = 1.00
    x = Node_df['พิกัด Y (m)  (.2float)'].tolist()
    x = list(set(x))
    x.sort()
    y = Node_df['พิกัด X (m)(.2float)'].tolist()
    y = list(set(y))
    y.sort()

    offset_x = (max(y)- min(y))/2
    offset_y = (max(x)- min(x))/2
    print(offset_x*2)
    print(offset_y*2)
    #แปลนคาน
    BeamDraw(obj = c,Beam_df= Beam_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    PltNode(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    DimDraw(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    ScaleTextDraw(obj = c,scale =scale,text='Beam Plan')


    #แปลน pointload
    c.showPage()
    PltNode(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    DimDraw(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    PointLoadDraw(obj = c,Reaction_df= Reaction_df,Node_df =  Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    ScaleTextDraw(obj = c,scale =scale,text='Node Plan')

    #แปลน distribute load
    c.showPage()
    BeamDraw(obj = c,Beam_df= Beam_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    PltNode(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    DimDraw(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    LinloadDraw(obj = c,LineLoad_df = LineLoad_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    ScaleTextDraw(obj = c,scale =scale,text= '')
    c.showPage()
    DimDraw(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    SlabDraw(obj = c,Slab_Data_df=Slab_Data_df,Slab_Result_df=Slab_Result_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    PltNode(obj = c,Node_df = Node_df,offset_x=offset_x,offset_y=offset_y,scale=scale)
    ScaleTextDraw(obj = c,scale =scale,text= 'Slab Plan')
    # บันทึกและปิดไฟล์
    c.save()

    print(f"สร้างไฟล์ PDF: {pdf_filepath}")
    """