import pandas as pd 
import os


def save_excel_sheet(df, filepath, sheetname, index=False):
    # Create file if it does not exist
    if not os.path.exists(filepath):
        df.to_excel(filepath, sheet_name=sheetname, index=index)

    # Otherwise, add a sheet. Overwrite if there exists one with the same name.
    else:
        with pd.ExcelWriter(filepath, engine='openpyxl', if_sheet_exists='replace', mode='a') as writer:
            df.to_excel(writer, sheet_name=sheetname, index=index)

def AllSlabLoadDistributed(filepath,DL,LL) :
    LoadCombine = pd.DataFrame({'DL': [DL],'LL' : [LL]})
    Slab_Data_df = pd.read_excel(filepath,sheet_name='Slab_Data')
    Node_df = pd.read_excel(filepath,sheet_name='Node_Data')
    SlabLoadCal_df = pd.DataFrame({'Slab No.': [],'Node1' : [],'Node2' : [],'linload1 (T/m)' : [] ,'Node3' : [],'Node4' : [],'linload2 (T/m)' : [],'Node5' : [],'Node6' : [],'linload3 (T/m)' : [],'Node7' : [],'Node8' : [],'linload4 (T/m)' : [],'Type' : [] })
    thickness_df = pd.read_excel(filepath.replace('.xlsx','-Analysed.xlsx'),sheet_name='Slab_Result')

    for index,value in enumerate(Slab_Data_df['หมายเลขพื้น']) :
        Type = str(Slab_Data_df.loc[index,'ประเภทแผ่นพื้น'])
        if Type == 'พื้นหล่อในที่' :
            for v , row_thickness in thickness_df.iterrows() :
                if row_thickness['หมายเลขพื้น'] == Slab_Data_df.loc[index,'หมายเลขพื้น'] :
                    thickness = float(row_thickness['Thickness'])
                    break
            #ตรวจสอบว่าเป็นพื้นสองทางหรือพื้นทางเดียว
            Ix = float(Slab_Data_df.loc[index,'Ix']) ; Jx = float(Slab_Data_df.loc[index,'Jx']) ; Iy = float(Slab_Data_df.loc[index,'Iy']) ; Ly = float(Slab_Data_df.loc[index,'Ly'])
            Xlenght = Jx-Ix ; Ylenght = Ly-Iy ; m = Xlenght / Ylenght
            if m > 2.0 or m < 0.5 : # พื้นทางเดียว
                if Xlenght > Ylenght : #แกน X ยาวกว่า ถ่ายแรงเหมือนพื้นสำเร็จวางแกน Y
                    #น้ำหนักพื้นถ่ายไปที่ Node I,J  กับ K,L Node1,Node2  Node5,Node6 ,linload1 linload3
                    Node1 = int(Slab_Data_df.loc[index,'จุดต่อ I']) ; Node2 =  int(Slab_Data_df.loc[index,'จุดต่อ J']) ; Node5 = int(Slab_Data_df.loc[index,'จุดต่อ K']) ; Node6 = int(Slab_Data_df.loc[index,'จุดต่อ L'])
                    Iy = float(Slab_Data_df.loc[index,'Iy']) ; Ly = float(Slab_Data_df.loc[index,'Ly']) ; lenght = (Ly - Iy)*0.5
                    #linload = DL + LL
                    linload = float(((thickness*2.4)+float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกคงที่อื่นๆ']))*float(LoadCombine.loc[0,'DL'])*lenght) + (float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกจรอื่นๆ'])*float(LoadCombine.loc[0,'LL'])*lenght)
                    new_row = pd.Series({'Slab No.': value ,'Node1' : Node1 ,'Node2' : Node2 ,'linload1 (T/m)' : linload ,'Node3' : 0 ,'Node4' : 0,'linload2 (T/m)' : 0,'Node5' : Node5,'Node6' : Node6 ,'linload3 (T/m)' : linload,'Node7' : 0,'Node8' : 0,'linload4 (T/m)' : 0,'Type' : 'พื้นทางเดียว' })
                    SlabLoadCal_df = pd.concat([SlabLoadCal_df, new_row.to_frame().T], ignore_index=True)                

                else : #แกน Y ยาวกว่า ถ่ายแรงเหมือนพื้นสำเร็จวางแกน X
                    #น้ำหนักพื้นถ่ายไปที่ Node L,I  กับ J,K Node1,Node8  Node5,Node6 ,linload1 linload3
                    Node3 = int(Slab_Data_df.loc[index,'จุดต่อ I']) ; Node4 =  int(Slab_Data_df.loc[index,'จุดต่อ L']) ; Node7 = int(Slab_Data_df.loc[index,'จุดต่อ J']) ; Node8 = int(Slab_Data_df.loc[index,'จุดต่อ K'])
                    Ix = float(Slab_Data_df.loc[index,'Ix']) ; Jx = float(Slab_Data_df.loc[index,'Jx']) ; lenght = (Jx - Ix)*0.5
                    #linload = DL + LL
                    linload = float(((thickness*2.4)+float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกคงที่อื่นๆ']))*float(LoadCombine.loc[0,'DL'])*lenght) + (float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกจรอื่นๆ'])*float(LoadCombine.loc[0,'LL'])*lenght)
                    new_row = pd.Series({'Slab No.': value ,'Node1' : 0 ,'Node2' : 0 ,'linload1 (T/m)' : 0,'Node3' : Node3 ,'Node4' : Node4,'linload2 (T/m)' : linload,'Node5' : 0,'Node6' : 0 ,'linload3 (T/m)' : 0,'Node7' : Node7,'Node8' : Node8,'linload4 (T/m)' : linload,'Type' : 'พื้นทางเดียว' })
                    SlabLoadCal_df = pd.concat([SlabLoadCal_df, new_row.to_frame().T], ignore_index=True)        
            else : #พื้นสองทาง
                if m == 1.0 : 
                    #thickness * (lenght/3) 
                    Ix = float(Slab_Data_df.loc[index,'Ix']) ; Jx = float(Slab_Data_df.loc[index,'Jx']) ; lenght = (Jx - Ix)
                    DL = (thickness*2.4*(lenght/3)) + float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกคงที่อื่นๆ']) ; LL = float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกจรอื่นๆ']) * (lenght/3)
                    linload = (DL+LL) 
                    new_row = pd.Series({'Slab No.': value ,'Node1' : int(Slab_Data_df.loc[index,'จุดต่อ I']) ,'Node2' : int(Slab_Data_df.loc[index,'จุดต่อ J']) ,'linload1 (T/m)' : linload ,'Node3' : int(Slab_Data_df.loc[index,'จุดต่อ J']) ,'Node4' : int(Slab_Data_df.loc[index,'จุดต่อ K']),'linload2 (T/m)' : linload,'Node5' : int(Slab_Data_df.loc[index,'จุดต่อ K']),'Node6' : int(Slab_Data_df.loc[index,'จุดต่อ L']) ,'linload3 (T/m)' : linload,'Node7' : int(Slab_Data_df.loc[index,'จุดต่อ L']),'Node8' : int(Slab_Data_df.loc[index,'จุดต่อ I']),'linload4 (T/m)' : linload,'Type' : 'พื้นสองทาง' })
                    SlabLoadCal_df = pd.concat([SlabLoadCal_df, new_row.to_frame().T], ignore_index=True)
                else :
                    if m > 1.0 : #ด้านนอนยาวกว่าด้านตั้ง
                        m = 1/m
                        #แรงด้านนอน (ด้านยาว)
                        factor = (3-(m**2))*0.5
                        Iy = float(Slab_Data_df.loc[index,'Iy']) ; Ly = float(Slab_Data_df.loc[index,'Ly']) ; lenght = (Ly - Iy)
                        DL = (thickness*2.4*(lenght/3)) + float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกคงที่อื่นๆ']) ; LL = (float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกจรอื่นๆ'])) * (lenght/3)
                        linload_long = (DL+LL) * factor  
                        linload_short = (DL+LL) 
                        new_row = pd.Series({'Slab No.': value ,'Node1' : int(Slab_Data_df.loc[index,'จุดต่อ I']) ,'Node2' : int(Slab_Data_df.loc[index,'จุดต่อ J']) ,'linload1 (T/m)' : linload_long ,'Node3' : int(Slab_Data_df.loc[index,'จุดต่อ J']) ,'Node4' : int(Slab_Data_df.loc[index,'จุดต่อ K']),'linload2 (T/m)' : linload_short,'Node5' : int(Slab_Data_df.loc[index,'จุดต่อ K']),'Node6' : int(Slab_Data_df.loc[index,'จุดต่อ L']) ,'linload3 (T/m)' : linload_long,'Node7' : int(Slab_Data_df.loc[index,'จุดต่อ L']),'Node8' : int(Slab_Data_df.loc[index,'จุดต่อ I']),'linload4 (T/m)' : linload_short,'Type' : 'พื้นสองทาง' })
                        SlabLoadCal_df = pd.concat([SlabLoadCal_df, new_row.to_frame().T], ignore_index=True)
                        pass
                    else : #ด้านตั้งยาวกว่าด้านนอน
                        factor = (3-(m**2))*0.5
                        Ix = float(Slab_Data_df.loc[index,'Ix']) ; Jx = float(Slab_Data_df.loc[index,'Jx']) ; lenght = (Jx - Ix)
                        DL = (thickness*2.4*(lenght/3)) ; LL = float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกจรอื่นๆ']) * (lenght/3)
                        linload_long = (DL+LL) * factor   
                        linload_short = DL+LL  
                        new_row = pd.Series({'Slab No.': value ,'Node1' : int(Slab_Data_df.loc[index,'จุดต่อ I']) ,'Node2' : int(Slab_Data_df.loc[index,'จุดต่อ J']) ,'linload1 (T/m)' : linload_short ,'Node3' : int(Slab_Data_df.loc[index,'จุดต่อ J']) ,'Node4' : int(Slab_Data_df.loc[index,'จุดต่อ K']),'linload2 (T/m)' : linload_long,'Node5' : int(Slab_Data_df.loc[index,'จุดต่อ K']),'Node6' : int(Slab_Data_df.loc[index,'จุดต่อ L']) ,'linload3 (T/m)' : linload_short,'Node7' : int(Slab_Data_df.loc[index,'จุดต่อ L']),'Node8' : int(Slab_Data_df.loc[index,'จุดต่อ I']),'linload4 (T/m)' : linload_long,'Type' : 'พื้นสองทาง' })
                        SlabLoadCal_df = pd.concat([SlabLoadCal_df, new_row.to_frame().T], ignore_index=True)


        elif Type == 'แผ่นพื้นสำเร็จวางแนวตั้ง' :
            #น้ำหนักพื้นถ่ายไปที่ Node I,J  กับ K,L Node1,Node2  Node5,Node6 ,linload1 linload3
            Node1 = int(Slab_Data_df.loc[index,'จุดต่อ I']) ; Node2 =  int(Slab_Data_df.loc[index,'จุดต่อ J']) ; Node5 = int(Slab_Data_df.loc[index,'จุดต่อ K']) ; Node6 = int(Slab_Data_df.loc[index,'จุดต่อ L'])
            Iy = float(Slab_Data_df.loc[index,'Iy']) ; Ly = float(Slab_Data_df.loc[index,'Ly']) ; lenght = (Ly - Iy)*0.5
            #linload = DL + LL
            linload = float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกคงที่อื่นๆ']*float(LoadCombine.loc[0,'DL'])*lenght) + (float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกจรอื่นๆ'])*float(LoadCombine.loc[0,'LL'])*lenght)
            new_row = pd.Series({'Slab No.': value ,'Node1' : Node1 ,'Node2' : Node2 ,'linload1 (T/m)' : linload ,'Node3' : 0 ,'Node4' : 0,'linload2 (T/m)' : 0,'Node5' : Node5,'Node6' : Node6 ,'linload3 (T/m)' : linload,'Node7' : 0,'Node8' : 0,'linload4 (T/m)' : 0,'Type' : 'แผ่นพื้นสำเร็จวางแนวตั้ง' })
            SlabLoadCal_df = pd.concat([SlabLoadCal_df, new_row.to_frame().T], ignore_index=True)
        elif Type == 'แผ่นพื้นสำเร็จวางแนวนอน' :
            #น้ำหนักพื้นถ่ายไปที่ Node L,I  กับ J,K Node1,Node8  Node5,Node6 ,linload1 linload3
            Node3 = int(Slab_Data_df.loc[index,'จุดต่อ I']) ; Node4 =  int(Slab_Data_df.loc[index,'จุดต่อ L']) ; Node7 = int(Slab_Data_df.loc[index,'จุดต่อ J']) ; Node8 = int(Slab_Data_df.loc[index,'จุดต่อ K'])
            Ix = float(Slab_Data_df.loc[index,'Ix']) ; Jx = float(Slab_Data_df.loc[index,'Jx']) ; lenght = (Jx - Ix)*0.5
            #linload = DL + LL
            linload = float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกคงที่อื่นๆ']*float(LoadCombine.loc[0,'DL'])*lenght) + (float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกจรอื่นๆ'])*float(LoadCombine.loc[0,'LL'])*lenght)
            new_row = pd.Series({'Slab No.': value ,'Node1' : 0 ,'Node2' : 0 ,'linload1 (T/m)' : 0,'Node3' : Node3 ,'Node4' : Node4,'linload2 (T/m)' : linload,'Node5' : 0,'Node6' : 0 ,'linload3 (T/m)' : 0,'Node7' : Node7,'Node8' : Node8,'linload4 (T/m)' : linload,'Type' : 'แผ่นพื้นสำเร็จวางแนวนอน' })
            SlabLoadCal_df = pd.concat([SlabLoadCal_df, new_row.to_frame().T], ignore_index=True)
            pass


    mergeSlabLoadDf = pd.DataFrame({'NodeList1' : [],'NodeList2' : [],'Load(T/M)' : [],'direction' : []})
    allSlabLoadDf = pd.DataFrame({'NodeList1' : [],'NodeList2' : [],'Load(T/M)' : [],'direction' : []})


    SlabLoadDf = SlabLoadCal_df
    for Row,i in enumerate(SlabLoadDf['Node1']) :
        node1 = int(SlabLoadDf.loc[Row,'Node1'])
        node2 = int(SlabLoadDf.loc[Row,'Node2'])
        for Rw,val in enumerate(Node_df['Node No. (int)']) :
            if val == node1 :
                pos1 = [float(Node_df.loc[Rw,'พิกัด X (m)(.2float)']),float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])]
            if val == node2 :
                pos2 = [float(Node_df.loc[Rw,'พิกัด X (m)(.2float)']),float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])]
        
        if pos1[0] == pos2[0] :
            direction = 'Y'
        elif pos1[1] == pos2[1] :
            direction = 'X'
        if i != 0 :
            if int(SlabLoadDf.loc[Row,'Node1']) > int(SlabLoadDf.loc[Row,'Node2']) :
                new_row = pd.Series({'NodeList1': int(SlabLoadDf.loc[Row,'Node2']) , 'NodeList2' : int(SlabLoadDf.loc[Row,'Node1']),'Load(T/M)' : float(SlabLoadDf.loc[Row,'linload1 (T/m)']),'direction' : direction})
            elif int(SlabLoadDf.loc[Row,'Node1']) < int(SlabLoadDf.loc[Row,'Node2']) :
                new_row = pd.Series({'NodeList1': int(SlabLoadDf.loc[Row,'Node1']) , 'NodeList2' : int(SlabLoadDf.loc[Row,'Node2']),'Load(T/M)' : float(SlabLoadDf.loc[Row,'linload1 (T/m)']),'direction' : direction})
            mergeSlabLoadDf = pd.concat([mergeSlabLoadDf, new_row.to_frame().T], ignore_index=True)
    for Row,i in enumerate(SlabLoadDf['Node3']) :
        node1 = int(SlabLoadDf.loc[Row,'Node3'])
        node2 = int(SlabLoadDf.loc[Row,'Node4'])
        for Rw,val in enumerate(Node_df['Node No. (int)']) :
            if val == node1 :
                pos1 = [float(Node_df.loc[Rw,'พิกัด X (m)(.2float)']),float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])]
            if val == node2 :
                pos2 = [float(Node_df.loc[Rw,'พิกัด X (m)(.2float)']),float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])]
        
        if pos1[0] == pos2[0] :
            direction = 'Y'
        elif pos1[1] == pos2[1] :
            direction = 'X'
        if i != 0 :
            if int(SlabLoadDf.loc[Row,'Node3']) > int(SlabLoadDf.loc[Row,'Node4']) :
                new_row = pd.Series({'NodeList1': int(SlabLoadDf.loc[Row,'Node4']) , 'NodeList2' : int(SlabLoadDf.loc[Row,'Node3']),'Load(T/M)' : float(SlabLoadDf.loc[Row,'linload2 (T/m)']),'direction' : direction})
            elif int(SlabLoadDf.loc[Row,'Node3']) < int(SlabLoadDf.loc[Row,'Node4']) :
                new_row = pd.Series({'NodeList1': int(SlabLoadDf.loc[Row,'Node3']) , 'NodeList2' : int(SlabLoadDf.loc[Row,'Node4']),'Load(T/M)' : float(SlabLoadDf.loc[Row,'linload2 (T/m)']),'direction' : direction})
            mergeSlabLoadDf = pd.concat([mergeSlabLoadDf, new_row.to_frame().T], ignore_index=True)




    for Row,i in enumerate(SlabLoadDf['Node5']) :
        node1 = int(SlabLoadDf.loc[Row,'Node5'])
        node2 = int(SlabLoadDf.loc[Row,'Node6'])
        for Rw,val in enumerate(Node_df['Node No. (int)']) :
            if val == node1 :
                pos1 = [float(Node_df.loc[Rw,'พิกัด X (m)(.2float)']),float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])]
            if val == node2 :
                pos2 = [float(Node_df.loc[Rw,'พิกัด X (m)(.2float)']),float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])]
        
        if pos1[0] == pos2[0] :
            direction = 'Y'
        elif pos1[1] == pos2[1] :
            direction = 'X'
        if i != 0 :
            if int(SlabLoadDf.loc[Row,'Node5']) > int(SlabLoadDf.loc[Row,'Node6']) :
                new_row = pd.Series({'NodeList1': int(SlabLoadDf.loc[Row,'Node6']) , 'NodeList2' : int(SlabLoadDf.loc[Row,'Node5']),'Load(T/M)' : float(SlabLoadDf.loc[Row,'linload3 (T/m)']),'direction' : direction})
            elif int(SlabLoadDf.loc[Row,'Node5']) < int(SlabLoadDf.loc[Row,'Node6']) :
                new_row = pd.Series({'NodeList1': int(SlabLoadDf.loc[Row,'Node5']) , 'NodeList2' : int(SlabLoadDf.loc[Row,'Node6']),'Load(T/M)' : float(SlabLoadDf.loc[Row,'linload3 (T/m)']),'direction' : direction})            
            mergeSlabLoadDf = pd.concat([mergeSlabLoadDf, new_row.to_frame().T], ignore_index=True)
    for Row,i in enumerate(SlabLoadDf['Node7']) :
        node1 = int(SlabLoadDf.loc[Row,'Node7'])
        node2 = int(SlabLoadDf.loc[Row,'Node8'])
        for Rw,val in enumerate(Node_df['Node No. (int)']) :
            if val == node1 :
                pos1 = [float(Node_df.loc[Rw,'พิกัด X (m)(.2float)']),float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])]
            if val == node2 :
                pos2 = [float(Node_df.loc[Rw,'พิกัด X (m)(.2float)']),float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])]
        
        if pos1[0] == pos2[0] :
            direction = 'Y'
        elif pos1[1] == pos2[1] :
            direction = 'X'
        if i != 0 :
            if int(SlabLoadDf.loc[Row,'Node7']) > int(SlabLoadDf.loc[Row,'Node8']) :
                new_row = pd.Series({'NodeList1': int(SlabLoadDf.loc[Row,'Node8']) , 'NodeList2' : int(SlabLoadDf.loc[Row,'Node7']),'Load(T/M)' : float(SlabLoadDf.loc[Row,'linload4 (T/m)']),'direction' : direction})
            elif int(SlabLoadDf.loc[Row,'Node7']) < int(SlabLoadDf.loc[Row,'Node8']) :
                new_row = pd.Series({'NodeList1': int(SlabLoadDf.loc[Row,'Node7']) , 'NodeList2' : int(SlabLoadDf.loc[Row,'Node8']),'Load(T/M)' : float(SlabLoadDf.loc[Row,'linload4 (T/m)']),'direction' : direction})
            
            mergeSlabLoadDf = pd.concat([mergeSlabLoadDf, new_row.to_frame().T], ignore_index=True)

    for Row,Value in enumerate(mergeSlabLoadDf['direction']) :
        node1 = int(mergeSlabLoadDf.loc[Row,'NodeList1'])
        node2 = int(mergeSlabLoadDf.loc[Row,'NodeList2'])
        Load = float(mergeSlabLoadDf.loc[Row,'Load(T/M)'])

        if Value == 'X' :
            for Rw,Node in enumerate(Node_df['Node No. (int)']) :
                if int(Node) == node1 : 
                    a = float(Node_df.loc[Rw,'พิกัด X (m)(.2float)'])
                    pos = float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])
                if int(Node) == node2 : b = float(Node_df.loc[Rw,'พิกัด X (m)(.2float)'])
            #print('a,b is :',a,b)
            List = [] 
            for Rw,Node in enumerate(Node_df['พิกัด X (m)(.2float)']) :
                Node = float(Node)
                if a <= Node <= b and float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'] == pos):
                    #print(a,'<=','Node is',Node,'<=',b)
                    List.append(Node_df.loc[Rw,'Node No. (int)'])
                    if len(List) == 2 :
                        addRowLoadSlab = pd.Series({'NodeList1' : List[0],'NodeList2' : List[1],'Load(T/M)' : Load,'direction' : 'X'})
                        #print(addRowLoadSlab)
                        allSlabLoadDf =  pd.concat([allSlabLoadDf, addRowLoadSlab.to_frame().T], ignore_index=True) 
                        #print(List)
                        List = [Node_df.loc[Rw,'Node No. (int)']]
                    
        elif Value == 'Y' :
            for Rw,Node in enumerate(Node_df['Node No. (int)']) :
                if int(Node) == node1 : 
                    a = float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])
                    pos = float(Node_df.loc[Rw,'พิกัด X (m)(.2float)'])
                if int(Node) == node2 : b = float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])
            #print('a,b is :',a,b)
            List = [] 
            for Rw,Node in enumerate(Node_df['พิกัด Y (m)  (.2float)']) :
                Node = float(Node)
                if a <= Node <= b and float(Node_df.loc[Rw,'พิกัด X (m)(.2float)'] == pos):
                    #print(a,'<=','Node is',Node,'<=',b)
                    List.append(Node_df.loc[Rw,'Node No. (int)'])
                    if len(List) == 2 :
                        addRowLoadSlab = pd.Series({'NodeList1' : List[0],'NodeList2' : List[1],'Load(T/M)' : Load,'direction' : 'Y'})
                        #print(addRowLoadSlab)
                        allSlabLoadDf =  pd.concat([allSlabLoadDf, addRowLoadSlab.to_frame().T], ignore_index=True) 
                        #print(List)
                        List = [Node_df.loc[Rw,'Node No. (int)']]
    save_excel_sheet(allSlabLoadDf,filepath,sheetname='allSlabLoadDistributed', index=False)                    
    return allSlabLoadDf




if __name__ == "__main__":
    AllSlabLoadDistributed(filepath = r'C:\ProjectVenv\.venv\ProjectHouse\Exhouse.xlsx')

