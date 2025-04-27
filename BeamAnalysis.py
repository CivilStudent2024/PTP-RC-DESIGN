from indeterminatebeam import Beam, Support, PointLoadV, DistributedLoadV
from BeamAstCaler import BeamAstcal,addStirrup
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


"""เสต็ปการทำงาน
1.BeamRank หาว่าคานตัวที่พิจารณาว่า มีคานตัวใดมาฝากมันไว้บ้าง
2.BeamCalculated ทำการลิสต์ลำดับการวิเคราะห์โครงสร้างว่าต้องวิเคราะห์คานหมายเลขใดก่อน
3.BeamCreated คือลิสต์ที่เก็บหมายเลขคานที่ทำการสร้างแล้ว
"""

def Analysis(filepath,fc,fy) :
    allSlabLoadDf = pd.read_excel(filepath,sheet_name='allSlabLoadDistributed')
    allLineLoadDf = pd.read_excel(filepath,sheet_name='allLineLoadDistributed')
    SectionTypeDf = pd.read_excel(filepath,sheet_name='SectionType')
    allBeamDf = pd.read_excel(filepath,sheet_name='allBeamDf')
    PointLoad_df = pd.read_excel(filepath,sheet_name='PointLoad_Data')
    Beam_df = pd.read_excel(filepath,sheet_name='Control_Parameter')
    reaction_df = pd.DataFrame({'Node': [], 'Load(T/M)': [] })
    Beam_data_df = pd.read_excel(filepath,sheet_name='Beam_Data')
    Control_df = pd.read_excel(filepath,sheet_name='Control_Parameter')
    for i,val in enumerate(PointLoad_df['ขนาดน้ำหนักบรรทุก (T/m)']) :
        if val != 0.0 :
            node = int(PointLoad_df.loc[i,'หมายเลขจุดต่อ'])
            new_row = pd.Series({'Node': node , 'Load(T/M)': -float(val)})
            reaction_df = pd.concat([reaction_df, new_row.to_frame().T], ignore_index=True)           


    BeamRank = {}
    for i in range(1,int(Beam_df.loc[0,'จำนวนคาน'])+1) :
        BeamRank[i] = []
    print(BeamRank)
    for index1,row1 in allBeamDf.iterrows() :
        if row1['ตำแหน่งNode2'] =='ใน' :
            FindNode = row1['Node List2']
            direction = row1['ทิศทางการวาง']
            for index2,row2 in allBeamDf.iterrows() :
                if row2['หมายเลขคาน'] != row1['หมายเลขคาน'] and row2['ทิศทางการวาง'] != direction :
                    if row2['Node List2'] == FindNode and row2['ตำแหน่งNode2'] == 'ใน':
                        #print('คานที่หาNodeที่ชน',row1['หมายเลขคาน'],'คานที่เจอNodeที่ชน',row2['หมายเลขคาน'],'Node ที่ชน',FindNode)
                        if row1['Node-List2 status'] == 'N' :
                            if row1['ทิศทางการวาง'] == 'X' :
                                BeamRank[row2['หมายเลขคาน']].append(row1['หมายเลขคาน'])
                            elif row1['ทิศทางการวาง'] == 'Y' :
                                BeamRank[row1['หมายเลขคาน']].append(row2['หมายเลขคาน'])
                        elif row1['Node-List2 status'] == 'E' :
                            if row1['ทิศทางการวาง'] == 'X' :
                                BeamRank[row1['หมายเลขคาน']].append(row2['หมายเลขคาน']) #สลับจาก row2,1 เป้น 1,2
                            elif row1['ทิศทางการวาง'] == 'Y' :
                                BeamRank[row2['หมายเลขคาน']].append(row1['หมายเลขคาน']) #สลับจาก row1,2 เป้น 2,1

                    elif row2['Node List2'] == FindNode and row2['ตำแหน่งNode2'] == 'นอก':
                        #print('คานที่หาNodeที่ชน',row1['หมายเลขคาน'],'คานที่เจอNodeที่ชน',row2['หมายเลขคาน'],'Node ที่ชน',FindNode)
                        if row1['Node-List2 status'] == 'N' :
                            BeamRank[row1['หมายเลขคาน']].append(row2['หมายเลขคาน'])
                        elif row1['Node-List2 status'] == 'E' :
                            BeamRank[row2['หมายเลขคาน']].append(row1['หมายเลขคาน'])
                    elif row2['Node List1'] == FindNode and row2['ตำแหน่งNode1'] == 'นอก':
                        #print('คานที่หาNodeที่ชน',row1['หมายเลขคาน'],'คานที่เจอNodeที่ชน',row2['หมายเลขคาน'],'Node ที่ชน',FindNode)
                        if row1['Node-List2 status'] == 'N' :
                            BeamRank[row1['หมายเลขคาน']].append(row2['หมายเลขคาน'])
                        elif row1['Node-List2 status'] == 'E' :
                            BeamRank[row2['หมายเลขคาน']].append(row1['หมายเลขคาน'])


        elif row1['ตำแหน่งNode2'] =='นอก' :
            FindNode = row1['Node List2']
            direction = row1['ทิศทางการวาง']
            for index2,row2 in allBeamDf.iterrows() :
                if row2['หมายเลขคาน'] != row1['หมายเลขคาน'] and row2['ทิศทางการวาง'] != direction :
                    if row2['Node List2'] == FindNode and row2['ตำแหน่งNode2'] == 'นอก':
                            if row1['Node-List2 status'] == 'N' :
                                if direction == 'X' :
                                    if row1['หมายเลขคาน'] not in BeamRank[row2['หมายเลขคาน']] :
                                        BeamRank[row2['หมายเลขคาน']].append(row1['หมายเลขคาน'])
                                else :
                                    if row2['หมายเลขคาน'] not in BeamRank[row1['หมายเลขคาน']] : 
                                        BeamRank[row1['หมายเลขคาน']].append(row2['หมายเลขคาน'])
                            elif row1['Node-List2 status'] == 'E' :
                                if direction == 'X' :
                                    BeamRank[row1['หมายเลขคาน']].append(row2['หมายเลขคาน'])
                                else : BeamRank[row2['หมายเลขคาน']].append(row1['หมายเลขคาน'])

                    elif row2['Node List1'] == FindNode and row2['ตำแหน่งNode1'] == 'นอก':
                            if row1['Node-List2 status'] == 'N' :
                                if direction == 'X' :
                                    if row1['หมายเลขคาน'] not in BeamRank[row2['หมายเลขคาน']] :
                                        BeamRank[row2['หมายเลขคาน']].append(row1['หมายเลขคาน'])
                                else : 
                                    if row2['หมายเลขคาน'] not in BeamRank[row1['หมายเลขคาน']] :
                                        BeamRank[row1['หมายเลขคาน']].append(row2['หมายเลขคาน'])
                            elif row1['Node-List2 status'] == 'E' :
                                if direction == 'X' :
                                    if row2['หมายเลขคาน'] not in BeamRank[row1['หมายเลขคาน']] :
                                        BeamRank[row1['หมายเลขคาน']].append(row2['หมายเลขคาน'])
                                else :
                                    if row1['หมายเลขคาน'] not in BeamRank[row2['หมายเลขคาน']] :
                                        BeamRank[row2['หมายเลขคาน']].append(row1['หมายเลขคาน'])    


        if row1['ตำแหน่งNode1'] =='นอก' :
            FindNode = row1['Node List1']
            direction = row1['ทิศทางการวาง']
            for index2,row2 in allBeamDf.iterrows() :
                if row2['หมายเลขคาน'] != row1['หมายเลขคาน'] and row2['ทิศทางการวาง'] != direction :
                    if row2['Node List2'] == FindNode and row2['ตำแหน่งNode2'] == 'นอก':
                            if row1['Node-List1 status'] == 'N' :
                                if direction == 'X' :
                                    if row1['หมายเลขคาน'] not in BeamRank[row2['หมายเลขคาน']] :
                                        BeamRank[row2['หมายเลขคาน']].append(row1['หมายเลขคาน'])
                                    
                                else : 
                                    if row2['หมายเลขคาน'] not in BeamRank[row1['หมายเลขคาน']] :
                                        BeamRank[row1['หมายเลขคาน']].append(row2['หมายเลขคาน'])
                            elif row1['Node-List1 status'] == 'E' :
                                if direction == 'X' :
                                    if row2['หมายเลขคาน'] not in BeamRank[row1['หมายเลขคาน']] : 
                                        BeamRank[row1['หมายเลขคาน']].append(row2['หมายเลขคาน'])
                                else :
                                    if row1['หมายเลขคาน'] not in BeamRank[row2['หมายเลขคาน']] :
                                        BeamRank[row2['หมายเลขคาน']].append(row1['หมายเลขคาน'])

                    elif row2['Node List1'] == FindNode and row2['ตำแหน่งNode1'] == 'นอก':
                            if row1['Node-List1 status'] == 'N' :
                                if direction == 'X' :
                                    if row1['หมายเลขคาน'] not in BeamRank[row2['หมายเลขคาน']] :
                                        BeamRank[row2['หมายเลขคาน']].append(row1['หมายเลขคาน'])
                                else : 
                                    if row2['หมายเลขคาน'] not in BeamRank[row1['หมายเลขคาน']] :
                                        BeamRank[row1['หมายเลขคาน']].append(row2['หมายเลขคาน'])
                            elif row1['Node-List1 status'] == 'E' :
                                if direction == 'X' :
                                    if row2['หมายเลขคาน'] not in BeamRank[row1['หมายเลขคาน']] : 
                                        BeamRank[row1['หมายเลขคาน']].append(row2['หมายเลขคาน'])
                                else :
                                    if row1['หมายเลขคาน'] not in BeamRank[row2['หมายเลขคาน']] : 
                                        BeamRank[row2['หมายเลขคาน']].append(row1['หมายเลขคาน'])    


    #นอกชนนอก (กำลังพัฒนา)






    # (สิ้นสุดนอกชนนอก)
    print(BeamRank)
    BeamNode = {}
    for i in BeamRank :
        BeamNode[i] = []
        for indexNode , row in allBeamDf.iterrows() :
            if int(row['หมายเลขคาน']) == i :
                if row['ตำแหน่งNode1'] == 'นอก' :
                    BeamNode[i].append(int(row['Node List1']))
                BeamNode[i].append(int(row['Node List2']))

    #จัดลำดับการวิเคราะห์
    print(BeamNode)
    BeamCalculated = [] 
    AllBeamDict = BeamRank
    loopcheck = 0
    while len(BeamCalculated) != len(AllBeamDict) :
        loopcheck += 1
        for i,val in enumerate(AllBeamDict) :
            if len(AllBeamDict[val]) == 0 and val not in BeamCalculated :
                BeamCalculated.append(val)
            elif len(AllBeamDict[val]) != 0 and val not in BeamCalculated :
                Check = False
                for j in AllBeamDict[val] :
                    if j in BeamCalculated :
                        Check = True
                    else : Check = False ; break
                if Check == True :
                    BeamCalculated.append(val)

        if loopcheck > 100 :
            print('กำหนดจุดคานไม่ถูกต้อง')
            print(BeamRank[7],BeamRank[12],BeamRank[13])
            break
    print(BeamCalculated)



    BeamCreated = []
    for beam in BeamCalculated :
        print('สร้างคาน',beam)
        for index,ebeam in enumerate(allBeamDf['หมายเลขคาน']) : #ebeam คือ element beam
            if ebeam == beam :
                if ebeam not in BeamCreated :
                    BeamCreated.append(ebeam)
                    #สร้างคาน
                    Lenght = 0
                    for index1 , row in allBeamDf.iterrows() :
                        if row['หมายเลขคาน'] == ebeam :
                            Lenght += row['Lspan']
                    globals()[f'beam{beam}'] = Beam(float(Lenght))
                    globals()[f'dataframe_beam{beam}'] = pd.DataFrame({'Span No':[],'Section List':[],'Position':[],'Moment':[],'Top-As':[],'Bot-As':[],'Shear':[],'RB6':[],'RB9':[],'Remark':[]})
                    globals()[f'beam{beam}'].update_units('force', 'T') ; globals()[f'beam{beam}'].update_units('distributed', 'T/m') ; globals()[f'beam{beam}'].update_units('moment', 'T.m') ; globals()[f'beam{beam}'].update_units('E', 'kg/cm2') ; globals()[f'beam{beam}'].update_units('I', 'mm4') ; globals()[f'beam{beam}'].update_units('deflection', 'mm') 
                    globals()[f'model_beam{beam}'] = {'Lenght':[],'supNode': [],'supPos' : [],'PLoad' : [],'PLoadNode' : [],'PLoadPos':[],'LLoad':[],'x0' : [],'x1' : []}  ###เพิ่มมา
                    globals()[f'model_beam{beam}']['Lenght'].append(float(Lenght)) ###เพิ่มมา
                    globals()[f'pos_support_beam{beam}'] = {}

                    #โค๊ดนี้ทำให้รันช้า

                    
                    distload = 0
                    
                    section_type = int(allBeamDf.loc[index,'TypeList'])
                    for rowType, type in enumerate(SectionTypeDf['หมายเลขหน้าตัดคาน']) :
                        if int(type) == section_type :
                            distload += -float((SectionTypeDf.loc[rowType,'SelfWeight']))
                    
                    if int(Control_df.loc[0,'จำนวนแผ่นพื้น']) != 0 :
                        for Count,Node in enumerate(allSlabLoadDf['NodeList1']) : #ถ่ายแรงจากพื้น
                            if Node == allBeamDf.loc[index,'Node List1'] and allSlabLoadDf.loc[Count,'NodeList2'] == allBeamDf.loc[index,'Node List2'] : 
                                distload += -float(allSlabLoadDf.loc[Count,'Load(T/M)'])

                    if int(Control_df.loc[0,'จำนวนน้ำหนักบรรทุกแบบกระจาย']) != 0 :
                        for Count,Node in enumerate(allLineLoadDf['NodeList1']) : #ถ่ายแรงจาก lineload
                            if Node == allBeamDf.loc[index,'Node List1'] and allLineLoadDf.loc[Count,'NodeList2'] == allBeamDf.loc[index,'Node List2'] : 
                                distload += -float(allLineLoadDf.loc[Count,'Load(T/M)'])
                    x0 = 0
                    x1 = float(allBeamDf.loc[index,'Lbeam'])
                    linload = DistributedLoadV(distload,(x0,x1))
                    if distload != 0 :
                        globals()[f'beam{beam}'].add_loads(linload)
                        globals()[f'model_beam{beam}']['LLoad'].append(distload) ; globals()[f'model_beam{beam}']['x0'].append(x0) ; globals()[f'model_beam{beam}']['x1'].append(x1) ###เพ่ิมมา




                    if allBeamDf.loc[index,'Node-List1 status'] == 'Y':
                        supNode = int(allBeamDf.loc[index,'Node List1'])
                        sup1 = Support(0,(1,1,0),nodeid = allBeamDf.loc[index,'Node List1'])
                        globals()[f'beam{beam}'].add_supports(sup1) ;  globals()[f'pos_support_beam{beam}'][supNode] = 0.0     
                        globals()[f'model_beam{beam}']['supNode'].append(supNode) ; globals()[f'model_beam{beam}']['supPos'].append(0.0) ###เพิ่มมา

                    else :
                            supNode = int(allBeamDf.loc[index,'Node List1'])
                            supportCheck = False # ตรวจสอบว่าเป็นคานฝากไหม เริ่มต้นไม่ม่ี ถ้าตรวจสอบแล้วมีจะเป็น True
                            for i in BeamRank :
                                if i != beam and beam in BeamRank[i] :
                                    for l in BeamNode :
                                        if i == l :
                                            if supNode in BeamNode[l] :
                                                supportCheck = True
                                                break
                            if supportCheck == True :
                                firstsup = Support(0,(1,1,0),nodeid = allBeamDf.loc[index,'Node List1'])
                                globals()[f'beam{beam}'].add_supports(firstsup) ;  globals()[f'pos_support_beam{beam}'][supNode] = 0.0
                                globals()[f'model_beam{beam}']['supNode'].append(supNode) ; globals()[f'model_beam{beam}']['supPos'].append(0.0) ###เพิ่มมา
                            elif supportCheck == False : #ไม่เป็นคานฝาก
                                force = 0
                                for rw,node in enumerate(reaction_df['Node']) : #ตรวจสอบว่าเป็นคานรับไหม
                                    if node == supNode :
                                        loadpos = 0
                                        force = float(reaction_df.loc[rw,'Load(T/M)'])                                   
                                if force != 0 :
                                    Pload = PointLoadV(force,0)
                                    globals()[f'beam{beam}'].add_loads(Pload)
                                    globals()[f'model_beam{beam}']['PLoad'].append(force) ; globals()[f'model_beam{beam}']['PLoadNode'].append(supNode) ; globals()[f'model_beam{beam}']['PLoadPos'].append(0.0) ###เพิ่มมา
                                    force = 0 
                
                        
                    if allBeamDf.loc[index,'ตำแหน่งNode2'] == 'นอก' :
                        supNode = int(allBeamDf.loc[index,'Node List2'])
                        if allBeamDf.loc[index,'Node-List2 status'] == 'Y' :
                            sup2 = Support(float(Lenght),(1,1,0),nodeid = allBeamDf.loc[index,'Node List2'])
                            globals()[f'beam{beam}'].add_supports(sup2) ; globals()[f'pos_support_beam{beam}'][supNode] = float(Lenght)
                            globals()[f'model_beam{beam}']['supNode'].append(supNode) ; globals()[f'model_beam{beam}']['supPos'].append(float(Lenght)) ###เพิ่มมา
                            print((x0,x1))
                            globals()[f'beam{beam}'].analyse()
                            for node in globals()[f'pos_support_beam{beam}'] :
                                re_pos = float(globals()[f'pos_support_beam{beam}'][node])
                                new_row = pd.Series({'Node': node , 'Load(T/M)': -globals()[f'beam{beam}'].get_reaction(re_pos)[1]})
                                reaction_df = pd.concat([reaction_df, new_row.to_frame().T], ignore_index=True)     
                        

                        else : # N or E                  
                            supNode = int(allBeamDf.loc[index,'Node List2'])
                            supportCheck = False # ตรวจสอบว่าเป็นคานฝากไหม เริ่มต้นไม่ม่ี ถ้าตรวจสอบแล้วมีจะเป็น True
                            for i in BeamRank :
                                if i != beam and beam in BeamRank[i] :
                                    for l in BeamNode :
                                        if i == l :
                                            if supNode in BeamNode[l] :
                                                supportCheck = True
                                                break
                            if supportCheck == True :
                                sup2 = Support(float(Lenght),(1,1,0),nodeid = allBeamDf.loc[index,'Node List2'])
                                globals()[f'beam{beam}'].add_supports(sup2) ; globals()[f'pos_support_beam{beam}'][supNode] = float(Lenght)
                                globals()[f'model_beam{beam}']['supNode'].append(supNode) ; globals()[f'model_beam{beam}']['supPos'].append(float(Lenght))
                            elif supportCheck == False : #ไม่เป็นคานฝาก
                                force = 0
                                for rw,node in enumerate(reaction_df['Node']) : #ตรวจสอบว่าเป็นคานรับไหม
                                    if node == supNode :
                                        loadpos = float(allBeamDf.loc[index,'Lbeam'])
                                        force += float(reaction_df.loc[rw,'Load(T/M)'])                                   
                                if force != 0 :
                                    Pload = PointLoadV(force,loadpos)
                                    globals()[f'beam{beam}'].add_loads(Pload)
                                    globals()[f'model_beam{beam}']['PLoad'].append(force) ; globals()[f'model_beam{beam}']['PLoadNode'].append(supNode) ; globals()[f'model_beam{beam}']['PLoadPos'].append(float(Lenght)) ###เพิ่มมา
                                    force = 0

                            globals()[f'beam{beam}'].analyse() 
                            for node in globals()[f'pos_support_beam{beam}'] :
                                re_pos = float(globals()[f'pos_support_beam{beam}'][node])
                                new_row = pd.Series({'Node': node , 'Load(T/M)': -globals()[f'beam{beam}'].get_reaction(re_pos)[1]})
                                reaction_df = pd.concat([reaction_df, new_row.to_frame().T], ignore_index=True)     
                            

                    #"""
                    elif allBeamDf.loc[index,'ตำแหน่งNode2'] == 'ใน' :
                        supNode = int(allBeamDf.loc[index,'Node List2'])
                        if allBeamDf.loc[index,'Node-List2 status'] == 'Y' :
                            subsuppos = float(allBeamDf.loc[index,'Lbeam'])
                            sup2 = Support(subsuppos,(1,1,0),nodeid = allBeamDf.loc[index,'Node List2'])
                            globals()[f'beam{beam}'].add_supports(sup2) ; globals()[f'pos_support_beam{beam}'][supNode] = subsuppos
                            globals()[f'model_beam{beam}']['supNode'].append(supNode) ; globals()[f'model_beam{beam}']['supPos'].append(subsuppos) ###เพิ่มมา
                        else : # N or E                  
                            supNode = int(allBeamDf.loc[index,'Node List2'])
                            supportCheck = False # ตรวจสอบว่าเป็นคานฝากไหม เริ่มต้นไม่ม่ี ถ้าตรวจสอบแล้วมีจะเป็น True
                            for i in BeamRank :
                                if i != beam and beam in BeamRank[i] :
                                    for l in BeamNode :
                                        if i == l :
                                            if supNode in BeamNode[l] :
                                                supportCheck = True
                                                break
                            if supportCheck == True :
                                subsuppos = float(allBeamDf.loc[index,'Lbeam'])
                                sup2 = Support(subsuppos,(1,1,0),nodeid = allBeamDf.loc[index,'Node List2'])
                                globals()[f'beam{beam}'].add_supports(sup2) ; globals()[f'pos_support_beam{beam}'][supNode] = subsuppos
                                globals()[f'model_beam{beam}']['supNode'].append(supNode) ; globals()[f'model_beam{beam}']['supPos'].append(subsuppos) ###เพิ่มมา
                            elif supportCheck == False : #ไม่เป็นคานฝาก
                                force = 0
                                for rw,node in enumerate(reaction_df['Node']) : #ตรวจสอบว่าเป็นคานรับไหม
                                    if node == supNode :
                                        loadpos = float(allBeamDf.loc[index,'Lbeam'])
                                        force += float(reaction_df.loc[rw,'Load(T/M)'])
                                if force != 0 :
                                    Pload = PointLoadV(force,loadpos)
                                    globals()[f'beam{beam}'].add_loads(Pload)
                                    globals()[f'model_beam{beam}']['PLoad'].append(force) ; globals()[f'model_beam{beam}']['PLoadNode'].append(supNode) ; globals()[f'model_beam{beam}']['PLoadPos'].append(loadpos) ###เพิ่มมา
                                    force = 0
                            #"""

                else :

                    #หา distload นี้ทำให้โปรแกรมรันช้า
                    
                    distload = 0

                    section_type = int(allBeamDf.loc[index,'TypeList'])

                    for rowType, type in enumerate(SectionTypeDf['หมายเลขหน้าตัดคาน']) :
                        if type == section_type :
                            distload += -float((SectionTypeDf.loc[rowType,'SelfWeight']))
                            
                    if int(Control_df.loc[0,'จำนวนแผ่นพื้น']) != 0 :
                        for Count,Node in enumerate(allSlabLoadDf['NodeList1']) : #ถ่ายแรงจากพื้น
                            if Node == allBeamDf.loc[index,'Node List1'] and allSlabLoadDf.loc[Count,'NodeList2'] == allBeamDf.loc[index,'Node List2'] : 
                                distload += -float(allSlabLoadDf.loc[Count,'Load(T/M)'])

                    if int(Control_df.loc[0,'จำนวนน้ำหนักบรรทุกแบบกระจาย']) != 0 :
                        for Count,Node in enumerate(allLineLoadDf['NodeList1']) : #ถ่ายแรงจาก lineload
                            if Node == allBeamDf.loc[index,'Node List1'] and allLineLoadDf.loc[Count,'NodeList2'] == allBeamDf.loc[index,'Node List2'] : 
                                distload += -float(allLineLoadDf.loc[Count,'Load(T/M)'])
                                
                    x0 = float(allBeamDf.loc[index,'Lbeam']) - float(allBeamDf.loc[index,'Lspan'])
                    x1 = float(allBeamDf.loc[index,'Lbeam'])
                    linload = DistributedLoadV(distload,(x0,x1))
                    if distload != 0 :
                        globals()[f'beam{beam}'].add_loads(linload)
                        globals()[f'model_beam{beam}']['LLoad'].append(distload) ; globals()[f'model_beam{beam}']['x0'].append(x0) ; globals()[f'model_beam{beam}']['x1'].append(x1) ###เพ่ิมมา
                    
                    if allBeamDf.loc[index,'ตำแหน่งNode2'] == 'นอก' : 
                        if allBeamDf.loc[index,'Node-List2 status'] == 'Y' :
                            supNode = int(allBeamDf.loc[index,'Node List2'])
                            lastsuppos = float(allBeamDf.loc[index,'Lbeam'])
                            lastsup = Support(lastsuppos,(1,1,0),nodeid = allBeamDf.loc[index,'Node List2'])
                            globals()[f'beam{beam}'].add_supports(lastsup) ; globals()[f'pos_support_beam{beam}'][supNode] = lastsuppos
                            globals()[f'model_beam{beam}']['supNode'].append(supNode) ; globals()[f'model_beam{beam}']['supPos'].append(lastsuppos) ###เพิ่มมา
                            globals()[f'beam{beam}'].analyse()
                            for node in globals()[f'pos_support_beam{beam}'] :
                                re_pos = float(globals()[f'pos_support_beam{beam}'][node])
                                new_row = pd.Series({'Node': node , 'Load(T/M)': -globals()[f'beam{beam}'].get_reaction(re_pos)[1]})
                                reaction_df = pd.concat([reaction_df, new_row.to_frame().T], ignore_index=True)     

                            
                        else :
                            supNode = int(allBeamDf.loc[index,'Node List2'])
                            supportCheck = False # ตรวจสอบว่าเป็นคานฝากไหม เริ่มต้นไม่ม่ี ถ้าตรวจสอบแล้วมีจะเป็น True
                            for i in BeamRank :
                                if i != beam and beam in BeamRank[i] :
                                    for l in BeamNode :
                                        if i == l :
                                            if supNode in BeamNode[l] :
                                                supportCheck = True
                                                break
                            if supportCheck == True :
                                lastsuppos = float(allBeamDf.loc[index,'Lbeam'])
                                lastsup = Support(lastsuppos,(1,1,0),nodeid = allBeamDf.loc[index,'Node List2'])
                                globals()[f'beam{beam}'].add_supports(lastsup) ; globals()[f'pos_support_beam{beam}'][supNode] = lastsuppos
                                globals()[f'model_beam{beam}']['supNode'].append(supNode) ; globals()[f'model_beam{beam}']['supPos'].append(lastsuppos) ###เพิ่มมา
                            elif supportCheck == False : #ไม่เป็นคานฝาก
                                force = 0
                                for rw,node in enumerate(reaction_df['Node']) : #ตรวจสอบว่าเป็นคานรับไหม
                                    if node == supNode :
                                        loadpos = float(allBeamDf.loc[index,'Lbeam'])
                                        force += float(reaction_df.loc[rw,'Load(T/M)'])
                                if force != 0 :
                                    Pload = PointLoadV(force,loadpos)
                                    globals()[f'beam{beam}'].add_loads(Pload)
                                    globals()[f'model_beam{beam}']['PLoad'].append(force) ; globals()[f'model_beam{beam}']['PLoadNode'].append(supNode) ; globals()[f'model_beam{beam}']['PLoadPos'].append(loadpos) ###เพิ่มมา

                                    force = 0
                            globals()[f'beam{beam}'].analyse()
                            for node in globals()[f'pos_support_beam{beam}'] :
                                re_pos = float(globals()[f'pos_support_beam{beam}'][node])
                                new_row = pd.Series({'Node': node , 'Load(T/M)': -globals()[f'beam{beam}'].get_reaction(re_pos)[1]})
                                reaction_df = pd.concat([reaction_df, new_row.to_frame().T], ignore_index=True)     

                            
                    else :
                        supNode = int(allBeamDf.loc[index,'Node List2'])
                        if allBeamDf.loc[index,'Node-List2 status'] == 'Y' :
                            subsuppos = float(allBeamDf.loc[index,'Lbeam'])
                            subsup = Support(subsuppos,(1,1,0),nodeid = allBeamDf.loc[index,'Node List2'])
                            globals()[f'beam{beam}'].add_supports(subsup) ; globals()[f'pos_support_beam{beam}'][supNode] = subsuppos
                            globals()[f'model_beam{beam}']['supNode'].append(supNode) ; globals()[f'model_beam{beam}']['supPos'].append(subsuppos) ###เพิ่มมา
                        else :
                            supNode = int(allBeamDf.loc[index,'Node List2'])
                            supportCheck = False # ตรวจสอบว่าเป็นคานฝากไหม เริ่มต้นไม่ม่ี ถ้าตรวจสอบแล้วมีจะเป็น True
                            for i in BeamRank :
                                if i != beam and beam in BeamRank[i] :
                                    for l in BeamNode :
                                        if i == l :
                                            if supNode in BeamNode[l] :
                                                supportCheck = True
                                                break

                            if supportCheck == True :
                                subsuppos = float(allBeamDf.loc[index,'Lbeam'])
                                subsup = Support(subsuppos,(1,1,0),nodeid = allBeamDf.loc[index,'Node List2'])
                                globals()[f'beam{beam}'].add_supports(subsup) ; globals()[f'pos_support_beam{beam}'][supNode] = subsuppos
                                globals()[f'model_beam{beam}']['supNode'].append(supNode) ; globals()[f'model_beam{beam}']['supPos'].append(subsuppos) ###เพิ่มมา
                            elif supportCheck == False : #ไม่เป็นคานฝาก
                                force = 0
                                for rw,node in enumerate(reaction_df['Node']) : #ตรวจสอบว่าเป็นคานรับไหม
                                    if node == supNode :
                                        loadpos = float(allBeamDf.loc[index,'Lbeam'])
                                        force += float(reaction_df.loc[rw,'Load(T/M)'])                                    
                                if force != 0 :
                                    Pload = PointLoadV(force,loadpos)
                                    globals()[f'beam{beam}'].add_loads(Pload)
                                    globals()[f'model_beam{beam}']['PLoad'].append(force) ; globals()[f'model_beam{beam}']['PLoadNode'].append(supNode) ; globals()[f'model_beam{beam}']['PLoadPos'].append(loadpos) ###เพิ่มมา
                                    force = 0

    df_save_path = filepath.replace('.xlsx','-Analysed.xlsx')
    print(reaction_df)
    reaction_df = reaction_df.groupby('Node',as_index=False).sum()
    print(reaction_df)
    reaction_result_df = pd.DataFrame({'Node' : [],'Load(T/M)' : []})
    for indexnode , s in enumerate(PointLoad_df['สถานะจุดต่อ']) :
        if s == 'Y' :
            for k ,rww in reaction_df.iterrows():
                if int(rww['Node']) == PointLoad_df.loc[indexnode,'หมายเลขจุดต่อ'] :
                    newrow_reaction_df = pd.Series({'Node': int(rww['Node']) , 'Load(T/M)': -float(rww['Load(T/M)'])})
                    reaction_result_df = pd.concat([reaction_result_df, newrow_reaction_df.to_frame().T], ignore_index=True)   



    save_excel_sheet(reaction_result_df,df_save_path,sheetname='Reaction_result',index=False)
    for i in range(1,int(Beam_df.loc[0,'จำนวนคาน'])+1) :
        model_df = pd.DataFrame({'Lenght':[],'supNode': [],'supPos' : [],'PLoad' : [],'PLoadNode' : [],'PLoadPos':[],'LLoad':[],'x0' : [],'x1' : []})
        data = globals()[f'model_beam{i}']
        df = pd.DataFrame.from_dict(data, orient='index').transpose()
        save_excel_sheet(df, df_save_path, sheetname=f'MB{i}', index=False)
        print("บันทึกไฟล์แล้ว")


    for x in range(1,int(Beam_df.loc[0,'จำนวนคาน'])+1) :
        fc=fc
        fy=fy

        for indexbeam, infobeam in enumerate(Beam_data_df['หมายเลขคาน']) :
            if infobeam == x :
                section_to_find = int(Beam_data_df.loc[indexbeam,'หน้าตัดคาน'])
                break
        for indexsection,infosection in enumerate(SectionTypeDf['หมายเลขหน้าตัดคาน']) :
            if infosection == section_to_find :
                width = float(SectionTypeDf.loc[indexsection,'ความกว้าง (m)'])
                depth = float(SectionTypeDf.loc[indexsection,'ความลึก (m)'])
                topcover = float(SectionTypeDf.loc[indexsection,'ระยะหุ้มเหล็กเสริมบน (m)'])
                botcover = float(SectionTypeDf.loc[indexsection,'ระยะหุ้มเหล็กเสริมล่าง (m)'])
        beam_list = []
        for i,key in enumerate(globals()[f'pos_support_beam{x}']) :
            beam_list.append(globals()[f'pos_support_beam{x}'][key])
        print(beam_list)

        for i in range((len(beam_list)-1)*5) :
            globals()[f'dataframe_beam{x}'].loc[i,'Span No'] = ''


        noSpan = 0 #เริ่มต้น span
        spanlenght = 0
        Lenght = float(globals()[f'model_beam{x}']['Lenght'][0])
        leftcantilevercheck = 0
        for index in range(len(beam_list)) :
            if index != len(beam_list) - 1 :
                if index == 0 and beam_list[index] != 0 :
                    position = 0
                    globals()[f'dataframe_beam{x}'].loc[5*index,'Position'] = position
                    globals()[f'dataframe_beam{x}'].loc[5*index,'Moment'] = globals()[f'beam{x}'].get_bending_moment(position)
                    globals()[f'dataframe_beam{x}'].loc[5*index,'Shear'] = globals()[f'beam{x}'].get_shear_force(position)
                    
                    Moment = float(globals()[f'dataframe_beam{x}'].loc[5*index,'Moment'])
                    try :
                        As = BeamAstcal(Mu = Moment ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100,topcover = topcover*100,botcover = botcover*100 ,safetyfactor=0.9) ; As.analyse()
                        As.get_As(a='bottom')
                        globals()[f'dataframe_beam{x}'].loc[(5*index),'Top-As'] = float(As.get_As(a='top'))
                        globals()[f'dataframe_beam{x}'].loc[(5*index),'Bot-As'] = float(As.get_As(a='bottom'))
                    except : 
                        globals()[f'dataframe_beam{x}'].loc[5*index,'Remark'] = 'Moment'
                        globals()[f'dataframe_beam{x}'].loc[(5*index),'Top-As'] = '-'
                        globals()[f'dataframe_beam{x}'].loc[(5*index),'Bot-As'] = '-'                        

                    shear = float(globals()[f'dataframe_beam{x}'].loc[5*index,'Shear'])
                    globals()[f'dataframe_beam{x}'].loc[(5*index),'RB6'] = addStirrup(Vu = shear ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100 ,botcover = botcover*100, safetyfactor = 0.85 ,stirType = 'RB6')
                    globals()[f'dataframe_beam{x}'].loc[(5*index),'RB9'] = addStirrup(Vu = shear ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100 ,botcover = botcover*100, safetyfactor = 0.85 ,stirType = 'RB9')

                    spanlenght = beam_list[index]
                    
                    globals()[f'dataframe_beam{x}'].loc[5*index,'Span No'] = f'คานยื่นทางซ้าย'
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+1,'Span No'] =f'คานหน้าตัดขนาด {width} x {depth}'
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+2,'Span No'] =f'ระยะหุ้มเหล็กเสริม {topcover} x {botcover}'
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+3,'Span No'] =f'ความยาว {round(spanlenght,3)} m'
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+4,'Span No'] = '---------------'
                    

                    for j in range(1,5) :
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Position'] = globals()[f'dataframe_beam{x}'].loc[5*index,'Position'] + ((j/4)*spanlenght)
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Moment'] = globals()[f'beam{x}'].get_bending_moment(globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Position'])
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Shear'] = globals()[f'beam{x}'].get_shear_force(globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Position'])

                        Moment = float(globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Moment'] )
                        try :
                            As = BeamAstcal(Mu = Moment ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100,topcover = topcover*100,botcover = botcover*100 ,safetyfactor=0.9) ; As.analyse()
                            As.get_As(a='bottom')
                            globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Top-As'] = float(As.get_As(a='top'))
                            globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Bot-As'] = float(As.get_As(a='bottom'))
                        except :
                            globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Remark'] = 'Moment'
                            globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Top-As'] = '-'
                            globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Bot-As'] = '-'

                        shear = float(globals()[f'dataframe_beam{x}'].loc[5*index+j,'Shear'])
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'RB6'] = addStirrup(Vu = shear ,fc = fc ,fy = fy ,width = width *100 ,depth = depth*100 ,botcover = botcover *100, safetyfactor = 0.85 ,stirType = 'RB6')
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'RB9'] = addStirrup(Vu = shear ,fc = fc ,fy = fy ,width = width *100,depth = depth*100 ,botcover = botcover *100, safetyfactor = 0.85 ,stirType = 'RB9')




                    #คอลัม section list
                    for j in range(0,5) :
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Section List'] = ((j/4)*spanlenght)
                    leftcantilevercheck = 1                


                noSpan += 1
                index += leftcantilevercheck
                if leftcantilevercheck == 0 :
                    position = beam_list[index]
                    spanlenght = spanlenght = beam_list[index+1]-beam_list[index]
                else : position = beam_list[index-1]
                globals()[f'dataframe_beam{x}'].loc[5*index,'Position'] = position
                globals()[f'dataframe_beam{x}'].loc[5*index,'Moment'] = globals()[f'beam{x}'].get_bending_moment(position)
                globals()[f'dataframe_beam{x}'].loc[5*index,'Shear'] = globals()[f'beam{x}'].get_shear_force(position)

                Moment = float(globals()[f'dataframe_beam{x}'].loc[5*index,'Moment'])
                try :
                    As = BeamAstcal(Mu = Moment ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100,topcover = topcover*100,botcover = botcover*100 ,safetyfactor=0.9) ; As.analyse()
                    As.get_As(a='bottom')
                    globals()[f'dataframe_beam{x}'].loc[(5*index),'Top-As'] = float(As.get_As(a='top'))
                    globals()[f'dataframe_beam{x}'].loc[(5*index),'Bot-As'] = float(As.get_As(a='bottom'))
                except : 
                    globals()[f'dataframe_beam{x}'].loc[(5*index),'Remark']  = 'Moment'    
                    globals()[f'dataframe_beam{x}'].loc[(5*index),'Top-As'] = '-'
                    globals()[f'dataframe_beam{x}'].loc[(5*index),'Bot-As'] = '-'

                shear = float(globals()[f'dataframe_beam{x}'].loc[5*index,'Shear'])
                globals()[f'dataframe_beam{x}'].loc[(5*index),'RB6'] = addStirrup(Vu = shear ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100 ,botcover = botcover*100, safetyfactor = 0.85 ,stirType = 'RB6')
                globals()[f'dataframe_beam{x}'].loc[(5*index),'RB9'] = addStirrup(Vu = shear ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100 ,botcover = botcover*100, safetyfactor = 0.85 ,stirType = 'RB9')


                globals()[f'dataframe_beam{x}'].loc[5*index,'Span No'] = f'ช่วงคานที่{noSpan}'
                globals()[f'dataframe_beam{x}'].loc[(5*index)+1,'Span No'] =f'คานหน้าตัดขนาด {width} x {depth}'
                globals()[f'dataframe_beam{x}'].loc[(5*index)+2,'Span No'] =f'ระยะหุ้มเหล็กเสริม {topcover} x {botcover}'
                globals()[f'dataframe_beam{x}'].loc[(5*index)+3,'Span No'] =f'ความยาว {round(spanlenght,3)} m'
                globals()[f'dataframe_beam{x}'].loc[(5*index)+4,'Span No'] = '---------------'

                if leftcantilevercheck == 1 :
                    spanlenght = beam_list[index]-beam_list[index-1]
                else : spanlenght = beam_list[index+1]-beam_list[index]

                for j in range(1,5) :
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Position'] = globals()[f'dataframe_beam{x}'].loc[5*index,'Position'] + ((j/4)*spanlenght)
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Moment'] = globals()[f'beam{x}'].get_bending_moment(globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Position'])
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Shear'] = globals()[f'beam{x}'].get_shear_force(globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Position'])
                    Moment = float(globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Moment'] )
                    try :
                        As = BeamAstcal(Mu = Moment ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100,topcover = topcover*100,botcover = botcover*100 ,safetyfactor=0.9) ; As.analyse()
                        As.get_As(a='bottom')
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Top-As'] = float(As.get_As(a='top'))
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Bot-As'] = float(As.get_As(a='bottom'))
                    except :
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Remark'] = 'Moment'
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Top-As'] = '-'
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Bot-As'] = '-'                    
                    shear = float(globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Shear'])
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'RB6'] = addStirrup(Vu = shear ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100 ,botcover = botcover*100, safetyfactor = 0.85 ,stirType = 'RB6')
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'RB9'] = addStirrup(Vu = shear ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100 ,botcover = botcover*100, safetyfactor = 0.85 ,stirType = 'RB9')
                #คอลัม section list
                for j in range(0,5) :
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Section List'] = ((j/4)*spanlenght)            

            else :
                if beam_list[index] != Lenght :
                    if leftcantilevercheck == 0 :
                        position = beam_list[index]
                    else : position = beam_list[index-1]
                    globals()[f'dataframe_beam{x}'].loc[5*index,'Position'] = position
                    globals()[f'dataframe_beam{x}'].loc[5*index,'Moment'] = globals()[f'beam{x}'].get_bending_moment(position)
                    globals()[f'dataframe_beam{x}'].loc[5*index,'Shear'] = globals()[f'beam{x}'].get_shear_force(position)

                    Moment = float(globals()[f'dataframe_beam{x}'].loc[5*index,'Moment'])
                    try :
                        As = BeamAstcal(Mu = Moment ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100,topcover = topcover*100,botcover = botcover*100 ,safetyfactor=0.9) ; As.analyse()
                        As.get_As(a='bottom')
                        globals()[f'dataframe_beam{x}'].loc[(5*index),'Top-As'] = float(As.get_As(a='top'))
                        globals()[f'dataframe_beam{x}'].loc[(5*index),'Bot-As'] = float(As.get_As(a='bottom'))
                    except :
                        globals()[f'dataframe_beam{x}'].loc[(5*index),'Remark'] = 'Moment'
                        globals()[f'dataframe_beam{x}'].loc[(5*index),'Top-As'] = '-'
                        globals()[f'dataframe_beam{x}'].loc[(5*index),'Bot-As'] = '-'                        


                    shear = float(globals()[f'dataframe_beam{x}'].loc[5*index,'Shear'])
                    globals()[f'dataframe_beam{x}'].loc[(5*index),'RB6'] = addStirrup(Vu = shear ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100 ,botcover = botcover*100, safetyfactor = 0.85 ,stirType = 'RB6')
                    globals()[f'dataframe_beam{x}'].loc[(5*index),'RB9'] = addStirrup(Vu = shear ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100 ,botcover = botcover*100, safetyfactor = 0.85 ,stirType = 'RB9')

                    spanlenght = Lenght-position
                    
                    globals()[f'dataframe_beam{x}'].loc[5*index,'Span No'] = f'คานยื่นทางขวา'
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+1,'Span No'] =f'คานหน้าตัดขนาด {width} x {depth}'
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+2,'Span No'] =f'ระยะหุ้มเหล็กเสริม {topcover} x {botcover}'
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+3,'Span No'] =f'ความยาว {round(spanlenght,3)} m'
                    globals()[f'dataframe_beam{x}'].loc[(5*index)+4,'Span No'] = '---------------'

                    for j in range(1,5) :
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Position'] = globals()[f'dataframe_beam{x}'].loc[5*index,'Position'] + ((j/4)*spanlenght)
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Moment'] = globals()[f'beam{x}'].get_bending_moment(globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Position'])
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Shear'] = globals()[f'beam{x}'].get_shear_force(globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Position'])

                        Moment = float(globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Moment'] )
                        try :
                            As = BeamAstcal(Mu = Moment ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100,topcover = topcover*100,botcover = botcover*100 ,safetyfactor=0.9) ; As.analyse()
                            As.get_As(a='bottom')
                            globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Top-As'] = float(As.get_As(a='top'))
                            globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Bot-As'] = float(As.get_As(a='bottom'))
                        except :
                            globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Remark'] = 'Moment'
                            globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Top-As'] = '-'
                            globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Bot-As'] = '-'

                        shear = float(globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Shear'])
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'RB6'] = addStirrup(Vu = shear ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100 ,botcover = botcover*100, safetyfactor = 0.85 ,stirType = 'RB6')
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'RB9'] = addStirrup(Vu = shear ,fc = fc ,fy = fy ,width = width*100 ,depth = depth*100 ,botcover = botcover*100, safetyfactor = 0.85 ,stirType = 'RB9')
                    
                    #คอลัม section list
                    for j in range(0,5) :
                        globals()[f'dataframe_beam{x}'].loc[(5*index)+j,'Section List'] = ((j/4)*spanlenght)

        save_excel_sheet(globals()[f'dataframe_beam{x}'], df_save_path, sheetname = f'Beam{x}', index=False)
        print(f'บันทึกผลการวิเคราะห์คาน {x} แล้ว')


    GB_list = {}
    GB_list_result = {}
    beam_add_check = []
    count_group = 0
    for i in range(1,int(Beam_df.loc[0,'จำนวนคาน'])+1) :
        if i not in beam_add_check :
            count_group += 1
            GB_list[f"B{count_group}"] = [f'B{i}']
            beam1 = pd.read_excel(df_save_path,sheet_name=f'Beam{i}') 
            beam1 = beam1[['Span No','Top-As','Bot-As','RB6','RB9']]
            for j in range(1,int(Beam_df.loc[0,'จำนวนคาน'])+1) :
                if j != i :
                    beam2 = pd.read_excel(df_save_path,sheet_name=f'Beam{j}')
                    beam2 = beam2[['Span No','Top-As','Bot-As','RB6','RB9']]
                    if  beam1.shape == beam2.shape and (beam1 == beam2).all().all() == True :
                        print('Group',i,j)
                        GB_list[f"B{count_group}"].append(f"B{j}")
                        beam_add_check.append(j)
            GB_list_result[f"B{count_group}"] = []            
            GB_list_result[f"B{count_group}"].append(','.join(GB_list[f"B{count_group}"])) 
    
    GB_list_result_df = pd.DataFrame({'กลุ่มคาน': [] , 'หมายเลขคาน': []})
    for i,val in enumerate(GB_list_result) :
        new_row = pd.Series({'กลุ่มคาน': f'{val}' , 'หมายเลขคาน': f'{GB_list_result[val]}'})
        GB_list_result_df = pd.concat([GB_list_result_df, new_row.to_frame().T], ignore_index=True)      
    save_excel_sheet(GB_list_result_df,df_save_path,sheetname='GroupBeam_List',index=False)

if __name__ == '__main__' :
    Analysis(filepath = r'C:\ProjectVenv\.venv\ProjectHouse\Exhouse.xlsx',fc=240,fy=4000)