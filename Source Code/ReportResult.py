from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle,Spacer,PageBreak,Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pandas as pd
import os

def Head_Write(filepath) :
    pdfmetrics.registerFont(TTFont('THSarabun', 'TH Sarabun New Regular.ttf'))
    styles = getSampleStyleSheet()
    thai_style = styles['Normal'].clone('ThaiStyle', fontName='THSarabun', fontSize=14)
    title_style = styles['Title'].clone('TitleThai', fontName='THSarabun', fontSize=14)
    head_df = pd.read_excel(filepath,sheet_name='Head_Title')



    project_title = head_df.loc[0,'Project :']
    floor_layer = head_df.loc[0,'Floor Layer :']
    engineer = head_df.loc[0,'Engineer :']
    date = head_df.loc[0,'Date :']

    head1  = Paragraph(f"Project : {project_title} {'&nbsp;'*65} Engineer : {engineer}", thai_style)
    head2 = Paragraph(f"Floor Layer :  {floor_layer} {'&nbsp;'*65} Date : {date} ",thai_style)
    headline = Paragraph(f"{'_'*85}",thai_style)

    Head_List = [head1,head2,headline]
    return Head_List

def SlabResult_Write(filepath) :
    pdfmetrics.registerFont(TTFont('THSarabun', 'TH Sarabun New Regular.ttf'))
    styles = getSampleStyleSheet()
    thai_style = styles['Normal'].clone('ThaiStyle', fontName='THSarabun', fontSize=14)
    title_style = styles['Title'].clone('TitleThai', fontName='THSarabun', fontSize=14)


    title = Paragraph("ผลการวิเคราะห์พื้น", title_style)
    data = pd.read_excel(filepath.replace('.xlsx','-Analysed.xlsx'),sheet_name='Slab_Result')
    data1 = data[['หมายเลขพื้น','SlabType','Case','Xlength','Ylength','Thickness','Top Cover','Bot Cover']]
    data1['Ylength'] = data1['Ylength'].round(3)
    data1.fillna('',inplace=True)
    data_list = [data1.columns.to_list()] + data1.values.tolist()
    ST_Detail = data[['หมายเลขพื้น','AS#1','AS#2','AS#3','AS#4','AS#5','AS#6','Use ST#1','Use ST#2','Use ST#3','Use ST#4','Use ST#5','Use ST#6',]]
    ST_Detail.rename(columns={'หมายเลขพื้น': 'No'}, inplace=True)
    ST_Detail.fillna('',inplace=True)
    ST_Detail_List = [ST_Detail.columns.to_list()] + ST_Detail.values.tolist()
    table1 = Table(data_list)
    table1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # สีพื้นหลังของแถวแรก
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # สีข้อความในแถวแรก
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # การจัดข้อความให้อยู่กลาง
        ('FONTNAME', (0, 0), (-1, -1), 'THSarabun'),  # ฟอนต์ของแถวแรก
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # การเพิ่มระยะห่างใต้แถวแรก
        ('GRID', (0, 0), (-1, -1), 1, colors.black),  # เส้นกริดของตาราง
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # สีพื้นหลังของแถวอื่น ๆ
        ("FONTSIZE", (0, 0), (-1, 0), 14),  # ฟอนต์หัวตาราง 14
        ("FONTSIZE", (0, 1), (-1, -1), 12)
    ]))

    table2 = Table(ST_Detail_List)
    table2.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # สีพื้นหลังของแถวแรก
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # สีข้อความในแถวแรก
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # การจัดข้อความให้อยู่กลาง
        ('FONTNAME', (0, 0), (-1, -1), 'THSarabun'),  # ฟอนต์ของแถวแรก
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # การเพิ่มระยะห่างใต้แถวแรก
        ('GRID', (0, 0), (-1, -1), 1, colors.black),  # เส้นกริดของตาราง
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # สีพื้นหลังของแถวอื่น ๆ
        ("FONTSIZE", (0, 0), (-1, 0), 14),  # ฟอนต์หัวตาราง 14
        ("FONTSIZE", (0, 1), (-1, -1), 12)
    ]))

    content = []
    [content.append(i) for i in Head_Write(filepath=filepath)]
    content.extend([Spacer(0,20),title,table1,Spacer(0,30),table2,PageBreak()])
    return content

def BeamResult_Write(filepath,nobeam) :
    pdfmetrics.registerFont(TTFont('THSarabun', 'TH Sarabun New Regular.ttf'))
    styles = getSampleStyleSheet()

    thai_style = styles['Normal'].clone('ThaiStyle', fontName='THSarabun', fontSize=14)
    title_style = styles['Title'].clone('TitleThai', fontName='THSarabun', fontSize=14)
    content = []
    for i in range(1,nobeam+1) :
        title = Paragraph(f"ผลการวิเคราะห์คานหมายเลข {i} ", title_style)
        img_path = os.getcwd()+f'\img_output\{i}.png'  
        img = Image(img_path, width=300, height=300,hAlign='CENTER') 
        data = pd.read_excel(filepath.replace('.xlsx','-Analysed.xlsx'),sheet_name=f'Beam{i}')
        data.fillna('',inplace=True)
        data[['Moment','Top-As','Bot-As','Section List','Position','Shear','RB6','RB9']] = data[['Moment','Top-As','Bot-As','Section List','Position','Shear','RB6','RB9']].round(3)
        if len(data) <= 15 :
            data_list = [data.columns.to_list()] + data.values.tolist()
            table1 = Table(data_list)
            table1.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # สีพื้นหลังของแถวแรก
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # สีข้อความในแถวแรก
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # การจัดข้อความให้อยู่กลาง
                ('FONTNAME', (0, 0), (-1, -1), 'THSarabun'),  # ฟอนต์ของแถวแรก
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # การเพิ่มระยะห่างใต้แถวแรก
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # เส้นกริดของตาราง
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # สีพื้นหลังของแถวอื่น ๆ
                ("FONTSIZE", (0, 0), (-1, 0), 14),  # ฟอนต์หัวตาราง 14
                ("FONTSIZE", (0, 1), (-1, -1), 12)
            ]))
            [content.append(i) for i in Head_Write(filepath=filepath)]
            content.extend([Spacer(0,20),title,img])
            content.extend([table1,PageBreak()])

        else :
            data1 = data.head(15)
            data2 = data.tail(len(data)-15)
            data1_list = [data1.columns.to_list()] + data1.values.tolist()
            data2_list = [data2.columns.to_list()] + data2.values.tolist()
            table1 = Table(data1_list)
            table1.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # สีพื้นหลังของแถวแรก
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # สีข้อความในแถวแรก
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # การจัดข้อความให้อยู่กลาง
                ('FONTNAME', (0, 0), (-1, -1), 'THSarabun'),  # ฟอนต์ของแถวแรก
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # การเพิ่มระยะห่างใต้แถวแรก
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # เส้นกริดของตาราง
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # สีพื้นหลังของแถวอื่น ๆ
                ("FONTSIZE", (0, 0), (-1, 0), 14),  # ฟอนต์หัวตาราง 14
                ("FONTSIZE", (0, 1), (-1, -1), 12)
            ]))
            table2 = Table(data2_list)
            table2.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # สีพื้นหลังของแถวแรก
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # สีข้อความในแถวแรก
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # การจัดข้อความให้อยู่กลาง
                ('FONTNAME', (0, 0), (-1, -1), 'THSarabun'),  # ฟอนต์ของแถวแรก
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # การเพิ่มระยะห่างใต้แถวแรก
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # เส้นกริดของตาราง
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # สีพื้นหลังของแถวอื่น ๆ
                ("FONTSIZE", (0, 0), (-1, 0), 14),  # ฟอนต์หัวตาราง 14
                ("FONTSIZE", (0, 1), (-1, -1), 12)
            ]))
            [content.append(i) for i in Head_Write(filepath=filepath)]
            content.extend([Spacer(0,20),title,img])
            content.extend([table1,PageBreak()])
            [content.append(i) for i in Head_Write(filepath=filepath)]
            content.extend([title,table2,PageBreak()])
    return content

def Material_Write(filepath,fc,fy,DL,LL) :
    pdfmetrics.registerFont(TTFont('THSarabun', 'TH Sarabun New Regular.ttf'))
    styles = getSampleStyleSheet()

    thai_style = styles['Normal'].clone('ThaiStyle', fontName='THSarabun', fontSize=14)
    title_style = styles['Title'].clone('TitleThai', fontName='THSarabun', fontSize=16)
    content = []
    title = Paragraph(f"Property of Material", title_style)
    line1 = Paragraph(f"{'&nbsp;'*20} fc\' = {fc} ksc ", thai_style)
    line2 = Paragraph(f"{'&nbsp;'*20}Wc = {2.400} T/m^3 ", thai_style)
    line3 = Paragraph(f"{'&nbsp;'*20}fy = {fy} ksc ", thai_style)
    line4 = Paragraph(f"{'&nbsp;'*20}Resistant M = {0.9} ", thai_style) 
    line5 = Paragraph(f"{'&nbsp;'*20}Resistant V = {0.85} ", thai_style) 
    line6 = Paragraph(f"{'&nbsp;'*20}Load Combination = {DL} DL + {LL} LL ", thai_style)    

    [content.append(i) for i in Head_Write(filepath=filepath)]
    content.extend([Spacer(0,20),title,line1,Spacer(0,10),line2,Spacer(0,10),line3,Spacer(0,10),line4,Spacer(0,10),line5,Spacer(0,10),line6,PageBreak()])
    return content

def Reaction_Write(filepath) :
    pdfmetrics.registerFont(TTFont('THSarabun', 'TH Sarabun New Regular.ttf'))
    styles = getSampleStyleSheet()

    thai_style = styles['Normal'].clone('ThaiStyle', fontName='THSarabun', fontSize=14)
    title_style = styles['Title'].clone('TitleThai', fontName='THSarabun', fontSize=14)
    content = []
    title = Paragraph(f"แรงปฏิกิริยาบนซัพพอร์ต", title_style)
    data = pd.read_excel(filepath.replace('.xlsx','-Analysed.xlsx'),sheet_name=f'Reaction_result')
    data['Load(T/M)'] = data['Load(T/M)'].round(3)
    data['Node'] = data['Node'].astype(str)
    data_list = [data.columns.to_list()] + data.values.tolist()
    table1 = Table(data_list)
    table1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # สีพื้นหลังของแถวแรก
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # สีข้อความในแถวแรก
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # การจัดข้อความให้อยู่กลาง
        ('FONTNAME', (0, 0), (-1, -1), 'THSarabun'),  # ฟอนต์ของแถวแรก
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # การเพิ่มระยะห่างใต้แถวแรก
        ('GRID', (0, 0), (-1, -1), 1, colors.black),  # เส้นกริดของตาราง
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # สีพื้นหลังของแถวอื่น ๆ
        ("FONTSIZE", (0, 0), (-1, 0), 14),  # ฟอนต์หัวตาราง 14
        ("FONTSIZE", (0, 1), (-1, -1), 12)
    ]))

    [content.append(i) for i in Head_Write(filepath=filepath)]
    content.extend([Spacer(0,20),title,table1,PageBreak()])    
    return content

def GroupBeam_Write(filepath) :
    pdfmetrics.registerFont(TTFont('THSarabun', 'TH Sarabun New Regular.ttf'))
    styles = getSampleStyleSheet()

    thai_style = styles['Normal'].clone('ThaiStyle', fontName='THSarabun', fontSize=14)
    title_style = styles['Title'].clone('TitleThai', fontName='THSarabun', fontSize=14)
    content = []
    title = Paragraph(f"การจัดเรียงกลุ่มคาน", title_style)
    data = pd.read_excel(filepath.replace('.xlsx','-Analysed.xlsx'),sheet_name=f'GroupBeam_List')
    data_list = [data.columns.to_list()] + data.values.tolist()
    table1 = Table(data_list)
    table1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # สีพื้นหลังของแถวแรก
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # สีข้อความในแถวแรก
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # การจัดข้อความให้อยู่กลาง
        ('FONTNAME', (0, 0), (-1, -1), 'THSarabun'),  # ฟอนต์ของแถวแรก
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # การเพิ่มระยะห่างใต้แถวแรก
        ('GRID', (0, 0), (-1, -1), 1, colors.black),  # เส้นกริดของตาราง
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # สีพื้นหลังของแถวอื่น ๆ
        ("FONTSIZE", (0, 0), (-1, 0), 14),  # ฟอนต์หัวตาราง 14
        ("FONTSIZE", (0, 1), (-1, -1), 12)
    ]))
    [content.append(i) for i in Head_Write(filepath=filepath)]
    content.extend([Spacer(0,20),title,table1,PageBreak()])    
    return content

def Report_Write(filepath,file_name,nobeam,fc,fy,DL,LL) :
    folder_name = "Report"
    os.makedirs(folder_name, exist_ok=True)
    report_save_path = os.getcwd() + '\Report'        
    document = SimpleDocTemplate(f'{report_save_path}\\{file_name}.pdf', pagesize=A4)
    elements = []
    [elements.append(i) for i in Material_Write(filepath=filepath,fc=fc,fy=fy,DL=DL,LL=LL)]
    [elements.append(i) for i in SlabResult_Write(filepath=filepath)]
    [elements.append(i) for i in GroupBeam_Write(filepath=filepath)]
    [elements.append(i) for i in BeamResult_Write(filepath=filepath,nobeam=nobeam)]
    [elements.append(i) for i in Reaction_Write(filepath=filepath)]
    document.build(elements)
    

if __name__ == '__main__' :
    file_name = "Deck.pdf"
    filepath = r'C:\ProjectVenv\.venv\ProjectHouse\Deck.xlsx'
    Report_Write(filepath=filepath,file_name="Report_Exhoue_Ex.pdf",nobeam=10)
    
    """
    document = SimpleDocTemplate(file_name, pagesize=A4)
    elements = []
    [elements.append(i) for i in Material_Write(filepath=filepath)]
    [elements.append(i) for i in SlabResult_Write(filepath=filepath)]
    [elements.append(i) for i in GroupBeam_Write(filepath=filepath)]
    [elements.append(i) for i in BeamResult_Write(filepath=filepath)]
    [elements.append(i) for i in Reaction_Write(filepath=filepath)]
    document.build(elements)
    """
