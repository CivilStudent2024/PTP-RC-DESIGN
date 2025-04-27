import sys
import os
import pandas as pd
from PyQt6.QtWidgets import QApplication, QMainWindow, QMenuBar, QMenu, QFileDialog, QMessageBox,QFrame,QLabel
from PyQt6.QtGui import QAction, QFont,QIcon
from PyQt6.QtCore import QSize
from HeadTitleWindowUI import TitleWindow 
from ControlDataWindowUI import ControlParameterWindow
from NodeDataWindowUI import NodeWindow
from SlabDataWindowUI import SlabDataWindow
from BeamDataWindowUI import BeamDataWindow
from PointLoadWindowUI import PointLoadWindow
from LineLoadWindowUI import LineLoadWindowui
from SectionTypeWindowUI import SectionTypeWindow
from plot_drawing import PlotWindow
from SlabLoadCal import AllSlabLoadDistributed
from allBeamDf import Allbeamdf
from LineLoadCal import allLineLoadcal
from BeamAnalysis import Analysis
from TwoWayAnalysis import TwoWay_analysed,OneWay_analysed,save_excel_sheet
from SlabResultUI import SlabResultWindow
from ReportResult import Report_Write
from PlanResult import Plan_Write
from ReactionResultUI import ReactionResultWindow

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # ตั้งชื่อหน้าต่าง
        self.setWindowTitle('PTP RC-DESIGN')
        self.setFixedSize(QSize(400,255))
        #OutFrame = QFrame(self) ; OutFrame.setFrameShape(QFrame.Shape.Box); OutFrame.setFrameShadow(QFrame.Shadow.Sunken) ; OutFrame.setLineWidth(3) 
        #OutFrame.setFixedSize(350,225);OutFrame.setStyleSheet('border:5px solid purple') ; OutFrame.move(25,50)
        frame1 = QFrame(self) ; frame1.setFrameShape(QFrame.Shape.Box); frame1.setFrameShadow(QFrame.Shadow.Sunken) ; frame1.setLineWidth(1) ;frame1.setFixedSize(217,16);frame1.setStyleSheet('background: "#800080"') ; frame1.move(0,16)
        frame2 = QFrame(self) ; frame2.setFrameShape(QFrame.Shape.Box); frame2.setFrameShadow(QFrame.Shadow.Sunken) ; frame2.setLineWidth(1) ;frame2.setFixedSize(325,16);frame2.setStyleSheet('background: "#D9D9D9"') ; frame2.move(75,48)
        frame3 = QFrame(self) ; frame3.setFrameShape(QFrame.Shape.Box); frame3.setFrameShadow(QFrame.Shadow.Sunken) ; frame3.setLineWidth(1) ;frame3.setFixedSize(400,16);frame3.setStyleSheet('background: "#800080"') ; frame3.move(0,80)
        frame4 = QFrame(self) ; frame4.setFrameShape(QFrame.Shape.Box); frame4.setFrameShadow(QFrame.Shadow.Sunken) ; frame4.setLineWidth(1) ;frame4.setFixedSize(400,16);frame4.setStyleSheet('background: "#800080"') ; frame4.move(0,175)
        frame5 = QFrame(self) ; frame5.setFrameShape(QFrame.Shape.Box); frame5.setFrameShadow(QFrame.Shadow.Sunken) ; frame5.setLineWidth(1) ;frame5.setFixedSize(325,16);frame5.setStyleSheet('background: "#D9D9D9"') ; frame5.move(75,207)
        frame6 = QFrame(self) ; frame6.setFrameShape(QFrame.Shape.Box); frame6.setFrameShadow(QFrame.Shadow.Sunken) ; frame6.setLineWidth(1) ;frame6.setFixedSize(217,16);frame6.setStyleSheet('background: "#800080"') ; frame6.move(0,239)
        self.lbTitle = QLabel('PTP',self) ;  font = QFont() ; font.setPointSize(20) ; self.lbTitle.setFont(font)  ; self.lbTitle.move(57,105) ; self.lbTitle.setFixedWidth(250) ; self.lbTitle.setStyleSheet('color : #E1FF00')
        self.lbTitle3 = QLabel('RC-DESIGN',self) ;  font = QFont() ; font.setPointSize(20) ; self.lbTitle3.setFont(font)  ; self.lbTitle3.move(120,105) ; self.lbTitle3.setFixedWidth(250) ; self.lbTitle3.setStyleSheet('color : #137B5A')
        self.lbTitle2 = QLabel('Version 1.00',self) ;  font = QFont() ; font.setPointSize(14) ; self.lbTitle2.setFont(font)  ; self.lbTitle2.move(57,135) ; self.lbTitle2.setFixedWidth(250) ; self.lbTitle2.setStyleSheet('color : #E3C9C9') 
        #self.lbTitle3 = QLabel('Version Demo ',self) ;  font = QFont() ; font.setPointSize(18) ; self.lbTitle3.setFont(font)  ; self.lbTitle3.move(130,180) ; self.lbTitle3.setFixedWidth(250)
        # สร้าง menubar
        menubar = self.menuBar()

        # สร้างเมนูต่างๆ
        file_menu = menubar.addMenu('File')
        edit_menu = menubar.addMenu('Edit')
        design_menu = menubar.addMenu('Design')
        printer_menu = menubar.addMenu('Printer')
        option_menu = menubar.addMenu('Option')

        # สร้าง action สำหรับเมนู File
        open_action = QAction('Open', self)
        open_action.triggered.connect(self.open_file_dialog)  # เชื่อมกับฟังก์ชันเปิดไฟล์
        file_menu.addAction(open_action)

        # สร้าง actions สำหรับเมนู Edit
        head_title_action = QAction('Head & Title', self)
        head_title_action.triggered.connect(self.head_title_action_triggered)


        control_data_action = QAction('Control Parameter', self)
        control_data_action.triggered.connect(self.control_data_action_triggered)

        node_data_action = QAction('Node Data', self)
        node_data_action.triggered.connect(self.node_data_action_triggered)

        slab_data_action = QAction('Slab Data', self)
        slab_data_action.triggered.connect(self.slab_data_action_triggered)

        section_data_action = QAction('Section Type', self)
        section_data_action.triggered.connect(self.section_data_action_triggered)

        beam_data_action = QAction('Beam Data', self)
        beam_data_action.triggered.connect(self.beam_data_action_triggered)

        PointLoad_data_action = QAction('Point Load', self)
        PointLoad_data_action.triggered.connect(self.PointLoad_data_action_triggered)

        LineLoad_data_action = QAction('Line Load', self)
        LineLoad_data_action.triggered.connect(self.LineLoad_data_action_triggered)

        # สร้าง actions สำหรับเมนู Design
        execute_action = QAction('Execute', self)
        execute_action.triggered.connect(self.execute_action_triggered)
        slab_result_action = QAction('Slab Result', self)
        slab_result_action.triggered.connect(self.slab_result_action_triggered)
        beam_result_action = QAction('Beam Result', self)
        beam_result_action.triggered.connect(self.beam_result_action_triggered)
        reaction_result_action = QAction('Reaction Result', self)
        reaction_result_action.triggered.connect(self.reaction_result_action_triggered)
        # สร้าง actions สำหรับเมนู Printer
        all_result_action = QAction('All Result', self)
        all_result_action.triggered.connect(self.all_result_action_triggered)
        # สร้าง actions สำหรับเมนู Option
        material_default_action = QMenu('Material Default', self)
        slab_default_action = QAction('Slab Default', self)
        beam_default_action = QAction('Beam Default', self)
        loadcom_menu = QMenu("Load Combination", self)
        fc_menu = QMenu("Concrete f\'c ", self)
        fy_menu = QMenu("Steel fy ", self)

        # เพิ่ม actions ในเมนู Edit
        edit_menu.addAction(head_title_action)
        edit_menu.addAction(control_data_action)
        edit_menu.addAction(node_data_action)
        edit_menu.addAction(slab_data_action)
        edit_menu.addAction(section_data_action)
        edit_menu.addAction(beam_data_action)
        edit_menu.addAction(PointLoad_data_action)
        edit_menu.addAction(LineLoad_data_action)


        # เพิ่ม actions ในเมนู Design
        design_menu.addAction(execute_action)
        design_menu.addAction(slab_result_action)
        design_menu.addAction(beam_result_action)
        design_menu.addAction(reaction_result_action)
        # เพิ่ม actions ในเมนู Printer
        printer_menu.addAction(all_result_action)

        # เพิ่ม actions ในเมนู Option
        option_menu.addMenu(material_default_action)
        #option_menu.addAction(slab_default_action)
        #option_menu.addAction(beam_default_action)
        option_menu.addMenu(loadcom_menu)

        material_default_action.addMenu(fc_menu)
        material_default_action.addMenu(fy_menu)


        self.DL = 1.4
        self.LL = 1.7
        self.loadcase1 = QAction("1.4 DL + 1.7 LL", self)
        self.loadcase1.setCheckable(True)
        self.loadcase1.setChecked(True)
        self.loadcase2 = QAction("1.2 DL + 1.6 LL", self)
        self.loadcase2.setCheckable(True)
        self.loadcase3 = QAction("1.0 DL + 1.0 LL", self)
        self.loadcase3.setCheckable(True)
        loadcom_menu.addAction(self.loadcase1)
        loadcom_menu.addAction(self.loadcase2)
        loadcom_menu.addAction(self.loadcase3)
        self.loadcase1.setObjectName('loadcase1')
        self.loadcase2.setObjectName('loadcase2')
        self.loadcase3.setObjectName('loadcase3')
        self.loadcase1.triggered.connect(self.loadcase_checkbox)
        self.loadcase2.triggered.connect(self.loadcase_checkbox)
        self.loadcase3.triggered.connect(self.loadcase_checkbox)


        self.fc = 210
        self.fc_case1 = QAction("210 ksc", self)
        self.fc_case1.setCheckable(True)
        self.fc_case1.setChecked(True)
        self.fc_case2 = QAction("240 ksc", self)
        self.fc_case2.setCheckable(True)
        self.fc_case3 = QAction("280 ksc", self)
        self.fc_case3.setCheckable(True)
        self.fc_case1.setObjectName('210')
        self.fc_case2.setObjectName('240')
        self.fc_case3.setObjectName('280')
        fc_menu.addAction(self.fc_case1)
        fc_menu.addAction(self.fc_case2)
        fc_menu.addAction(self.fc_case3)
        self.fc_case1.triggered.connect(self.fc_checkbox)
        self.fc_case2.triggered.connect(self.fc_checkbox)
        self.fc_case3.triggered.connect(self.fc_checkbox)

        self.fy = 4000
        self.fy_case1 = QAction("3000 ksc", self)
        self.fy_case1.setCheckable(True)
        self.fy_case2 = QAction("4000  ksc", self)
        self.fy_case2.setCheckable(True)
        self.fy_case2.setChecked(True)
        self.fy_case1.setObjectName('3000')
        self.fy_case2.setObjectName('4000')
        fy_menu.addAction(self.fy_case1)
        fy_menu.addAction(self.fy_case2)
        self.fy_case1.triggered.connect(self.fy_checkbox)
        self.fy_case2.triggered.connect(self.fy_checkbox)

        



        # กำหนดขนาดเริ่มต้นของหน้าต่าง
        self.new_window = None
        self.setGeometry(300, 300, 400, 300)

    # ฟังก์ชันที่เรียกใช้เมื่อคลิก action ต่างๆ

    def open_file_dialog(self):
        # เปิดไฟล์ดialog
        self.file_name, _ = QFileDialog.getOpenFileName(self, "Open File", "", "All Files (*);;Text Files (*.txt)")
        if self.file_name:
            QMessageBox.information(self, "File Selected", f"Selected file: {self.file_name}")
        else:
            QMessageBox.warning(self, "No File", "No file was selected!")
    def new_file_dialog(self) :
        pass

    # ฟังก์ชันสำหรับ action ของ Edit menu
    def head_title_action_triggered(self):
        self.open_new_window('Head & Title')
    def control_data_action_triggered(self):
        self.open_new_window('Control Parameter')
    def node_data_action_triggered(self):
        self.open_new_window('Node Data')
    def slab_data_action_triggered(self):
        self.open_new_window('Slab Data')
    def section_data_action_triggered(self):
        self.open_new_window('Section Type')         
    def beam_data_action_triggered(self):
        self.open_new_window('Beam Data')
    def PointLoad_data_action_triggered(self):
        self.open_new_window('Point Load')
    def LineLoad_data_action_triggered(self):
        self.open_new_window('Line Load')
    def execute_action_triggered(self):
        print('fy',self.fy,'f\'c',self.fc,'DL',self.DL,'LL',self.LL)
        filepath = self.file_name
        try :
            Control_df = pd.read_excel(filepath,sheet_name='Control_Parameter')
        except :
            pass
        
        self.show_info_message1()
        QMessageBox.information(self, "Parameter",f"f\'c = {self.fc}\nfy = {self.fy}\nslab min diameter bar : 9 mm\nLoad Combination : {self.DL} DL + {self.LL} LL") 
        df_save_path = filepath.replace('.xlsx','-Analysed.xlsx')
        if Control_df.loc[0,'จำนวนแผ่นพื้น']  != 0 :
            twoway_df = TwoWay_analysed(filepath=filepath,DL = self.DL,LL = self.LL,fc = self.fc,fy = self.fy)
            oneway_df = OneWay_analysed(filepath=filepath,DL = self.DL,LL = self.LL,fc = self.fc,fy = self.fy)
            all_slab = pd.concat([twoway_df,oneway_df],ignore_index = True)
            all_slab = all_slab.sort_values(by='หมายเลขพื้น')
            save_excel_sheet(all_slab, filepath = df_save_path , sheetname='Slab_Result', index=False)
            AllSlabLoadDistributed(filepath=filepath,DL = self.DL,LL = self.LL)
        if Control_df.loc[0,'จำนวนน้ำหนักบรรทุกแบบกระจาย'] != 0 :
            allLineLoadcal(filepath=filepath,DL = self.DL,LL = self.LL)

        Allbeamdf(filepath=filepath)
        Analysis(filepath=filepath,fc=self.fc,fy=self.fy)
        self.show_info_message2()        

    def all_result_action_triggered(self) :
        try :
            filepath = self.file_name
            Project_df = pd.read_excel(filepath,sheet_name='Head_Title')
            Control_df = pd.read_excel(filepath,sheet_name='Control_Parameter')
            nobeam = int(Control_df.loc[0,'จำนวนคาน'])
            Project_name = Project_df.loc[0,'Project :'] 
            Report_Write(filepath=filepath,file_name=Project_name,nobeam=nobeam,fc=self.fc,fy=self.fy,DL=self.DL,LL=self.LL)
            Plan_Write(filepath=filepath,file_name=Project_name)
            QMessageBox.information(self, "แจ้งเตือน","สร้างรายงานเสร็จเรียบร้อยแล้ว")

        except :
            pass
        



    def beam_result_action_triggered(self):
        self.open_new_window('Beam Result')
    def slab_result_action_triggered(self):
        self.open_new_window('Slab Result')
    def reaction_result_action_triggered(self):
        self.open_new_window('Reaction Result')
    def open_new_window(self,title):
        # เปิดหน้าต่างใหม่จากไฟล์ new_window.py
        if self.new_window is None:
            if title == 'Head & Title' :
                self.new_window = TitleWindow(filename=self.file_name)
            elif title == 'Control Parameter' :
                self.new_window = ControlParameterWindow(filename=self.file_name)
            elif title == 'Node Data' :
                self.new_window = NodeWindow(filename=self.file_name)                
            elif title == 'Slab Data' :
                self.new_window = SlabDataWindow(filename=self.file_name) 
            elif title == 'Section Type' :
                self.new_window = SectionTypeWindow(filename=self.file_name,DL = self.DL) 
            elif title == 'Beam Data' :
                self.new_window = BeamDataWindow(filename=self.file_name) 
            elif title == 'Point Load' :
                self.new_window = PointLoadWindow(filename=self.file_name) 
            elif title == 'Line Load' :
                self.new_window = LineLoadWindowui(filename=self.file_name) 
            elif title == 'Execute' :
                pass

            elif title == 'Beam Result' :
                try :
                    filepath =self.file_name.replace('.xlsx','-Analysed.xlsx')
                    self.show_info_message1()
                    self.new_window = PlotWindow(filepath=filepath,DL=self.DL,LL=self.LL)
                except : self.show_error_message1()
            elif title == 'Slab Result' :
                try :
                    filepath =self.file_name.replace('.xlsx','-Analysed.xlsx')
                    self.show_info_message1()
                    self.new_window = SlabResultWindow(filepath=filepath,DL=self.DL,LL=self.LL)
                except : self.show_error_message1()                
            elif title == 'Reaction Result' :
                try :
                    filepath =self.file_name.replace('.xlsx','-Analysed.xlsx')
                    self.show_info_message1()
                    self.new_window = ReactionResultWindow(filepath=filepath,DL=self.DL,LL=self.LL)
                except : self.show_error_message1()       
        self.new_window.window_closed.connect(self.on_window_closed)
        self.new_window.show()  # ต้องให้ new_window แสดง

    def loadcase_checkbox(self):
        sender = self.sender()
        if sender.objectName() == 'loadcase1' :
            self.loadcase1.setChecked(True)
            self.loadcase2.setChecked(False)
            self.loadcase3.setChecked(False)
            self.DL = 1.4
            self.LL = 1.7
        elif sender.objectName() == 'loadcase2' :
            self.loadcase2.setChecked(True)
            self.loadcase1.setChecked(False)
            self.loadcase3.setChecked(False)
            self.DL = 1.2
            self.LL = 1.6
        elif sender.objectName() == 'loadcase3' :
            self.loadcase3.setChecked(True)
            self.loadcase2.setChecked(False)
            self.loadcase1.setChecked(False)
            self.DL = 1.0
            self.LL = 1.0

    def fc_checkbox(self):
        sender = self.sender()
        if sender.objectName() == '210' :
            self.fc_case1.setChecked(True)
            self.fc_case2.setChecked(False)
            self.fc_case3.setChecked(False)
            self.fc = 210
        elif sender.objectName() == '240' :
            self.fc_case2.setChecked(True)
            self.fc_case1.setChecked(False)
            self.fc_case3.setChecked(False)
            self.fc = 240
        elif sender.objectName() == '280' :
            self.fc_case3.setChecked(True)
            self.fc_case1.setChecked(False)
            self.fc_case2.setChecked(False)
            self.fc = 280

    def fy_checkbox(self):
        sender = self.sender()
        if sender.objectName() == '3000' :
            self.fy_case1.setChecked(True)
            self.fy_case2.setChecked(False)
            self.fy = 3000
        elif sender.objectName() == '4000' :
            self.fy_case2.setChecked(True)
            self.fy_case1.setChecked(False)
            self.fy = 4000


    def on_window_closed(self) :
        print('New window was closed')
        self.new_window = None

    def closeEvent(self, event) :
        reply = QMessageBox.question(self,"ปิดหน้าต่าง","คุณแน่ใจไหมว่าต้องการออกจากโปรแกรม",QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes :
            event.accept()
        else:
            event.ignore()       

    def show_info_message1(self):
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Information) 
        msg_box.setWindowTitle("แจ้งเตือน")
        msg_box.setText("กำลังตรวจสอบและจัดเตรียม โปรดรอสักครู่")
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.setDefaultButton(QMessageBox.StandardButton.Ok)
        msg_box.exec()

    def show_info_message2(self):
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Information) 
        msg_box.setWindowTitle("แจ้งเตือน")
        msg_box.setText("วิเคราะห์เสร็จแล้ว")
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.setDefaultButton(QMessageBox.StandardButton.Ok)
        msg_box.exec()



    def show_error_message1(self):
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Critical) 
        msg_box.setWindowTitle("ข้อผิดพลาด")
        msg_box.setText("โปรดกรอกข้อมูลส่วนอื่นก่อนครับ")
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.setDefaultButton(QMessageBox.StandardButton.Ok)
        msg_box.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainpath = os.getcwd()
    mainpath = mainpath.replace('\\','/')
    with open(f'{mainpath}/AMOLED.qss') as style :
        app.setStyleSheet(style.read())
    #app.setWindowIcon(QIcon("logo.png"))
    window = MyWindow()
    window.show()
    sys.exit(app.exec())