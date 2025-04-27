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

def Allbeamdf(filepath) :
    filepath = filepath
    Node_df = pd.read_excel(filepath,sheet_name='Node_Data')
    Beam_df = pd.read_excel(filepath,sheet_name='Beam_Data')
    allBeamDf = pd.DataFrame({'หมายเลขคาน' : [] , 'Node List1' : [],'Node List2' : [],'Node1pos' : [],'Node2pos' : [] ,'Lspan' : [],'Lbeam' : [] ,'Node-List1 status' : [] , 'Node-List2 status' : [], 'ตำแหน่งNode1' : [], 'ตำแหน่งNode2' : [],'ทิศทางการวาง' : [],'TypeList' : []})
    allSlabLoadDf = pd.read_excel(filepath,sheet_name='allSlabLoadDistributed')
    for Row,Value in enumerate(Beam_df['ทิศทางการวาง']) :
        node1 = int(Beam_df.loc[Row,'Node List1']) 
        node2 = int(Beam_df.loc[Row,'Node List2']) 
        noBeam = int(Beam_df.loc[Row,'หมายเลขคาน']) 
        
        ListLspan = []
        TypeList = int(Beam_df.loc[Row,'หน้าตัดคาน'])
        if Value == 'X' :
            for Rw,Node in enumerate(Node_df['Node No. (int)']) :
                if int(Node) == node1 : 
                    a = float(Node_df.loc[Rw,'พิกัด X (m)(.2float)'])
                    pos = float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])
                if int(Node) == node2 : b = float(Node_df.loc[Rw,'พิกัด X (m)(.2float)'])
            #print('a,b is :',a,b)
            List = []
            ListStatus = [] 
            Listpos = []
            ListNodeType = []
            for Rw,Node in enumerate(Node_df['พิกัด X (m)(.2float)']) :
                Node = float(Node)
                if a <= Node <= b and float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)']) == pos:
                    #print(a,'<=','Node is',Node,'<=',b)
                    List.append(Node_df.loc[Rw,'Node No. (int)'])
                    ListStatus.append(Node_df.loc[Rw,'สถานะจุดต่อ'])
                    Listpos.append(Node_df.loc[Rw,'พิกัด X (m)(.2float)'])
                    if int(Node_df.loc[Rw,'Node No. (int)']) == node1 or int(Node_df.loc[Rw,'Node No. (int)']) == node2 :
                        ListNodeType.append('นอก')
                    else : ListNodeType.append('ใน')

                    if len(List) == 2 :
                        Lspan = Listpos[1]-Listpos[0]
                        ListLspan.append(Lspan)
                        
                        SlabLoad = 0
                        for Count,Node in enumerate(allSlabLoadDf['NodeList1']) : #ถ่ายแรงจากพื้น
                            if Node == List[0] and allSlabLoadDf.loc[Count,'NodeList2'] == List[1] : 
                                SlabLoad += float(allSlabLoadDf.loc[Count,'Load(T/M)'])                     



                        addRowBeam = pd.Series({'หมายเลขคาน' : noBeam ,'Node List1' : List[0],'Node List2' : List[1],'Node1pos' : Listpos[0],'Node2pos' : Listpos[1],'Lspan' : Listpos[1]-Listpos[0],'Lbeam' : sum(ListLspan)  ,'Node-List1 status' : ListStatus[0] , 'Node-List2 status' : ListStatus[1], 'ตำแหน่งNode1' : ListNodeType[0], 'ตำแหน่งNode2' : ListNodeType[1],'ทิศทางการวาง' : 'X','TypeList' : TypeList})
                        #print(addRowBeam)
                        allBeamDf =  pd.concat([allBeamDf, addRowBeam.to_frame().T], ignore_index=True) 
                        #print(List)
                        List = [Node_df.loc[Rw,'Node No. (int)']]
                        ListStatus = [Node_df.loc[Rw,'สถานะจุดต่อ']]
                        Listpos = [Node_df.loc[Rw,'พิกัด X (m)(.2float)']]
                        ListNodeType = ['ใน']

        elif Value == 'Y' :
            for Rw,Node in enumerate(Node_df['Node No. (int)']) :
                if int(Node) == node1 : 
                    a = float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])
                    pos = float(Node_df.loc[Rw,'พิกัด X (m)(.2float)'])
                if int(Node) == node2 : b = float(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])
            #print('a,b is :',a,b)
            List = []
            ListStatus = []
            Listpos = []
            ListNodeType = []



            for Rw,Node in enumerate(Node_df['พิกัด Y (m)  (.2float)']) :
                Node = float(Node)
                if a <= Node <= b and float(Node_df.loc[Rw,'พิกัด X (m)(.2float)'] == pos):
                    #print(a,'<=','Node is',Node,'<=',b)
                    List.append(Node_df.loc[Rw,'Node No. (int)'])
                    ListStatus.append(Node_df.loc[Rw,'สถานะจุดต่อ'])
                    Listpos.append(Node_df.loc[Rw,'พิกัด Y (m)  (.2float)'])
                            
                    if int(Node_df.loc[Rw,'Node No. (int)']) == node1 or int(Node_df.loc[Rw,'Node No. (int)']) == node2 :
                        ListNodeType.append('นอก')
                    else : ListNodeType.append('ใน')
            
                    if len(List) == 2 :
                        Lspan = Listpos[1]-Listpos[0]
                        ListLspan.append(Lspan)

                        addRowBeam = pd.Series({'หมายเลขคาน' : noBeam ,'Node List1' : List[0],'Node List2' : List[1],'Node1pos' : Listpos[0],'Node2pos' : Listpos[1],'Lspan' : Listpos[1]-Listpos[0],'Lbeam' : sum(ListLspan) , 'Node-List1 status' : ListStatus[0] , 'Node-List2 status' : ListStatus[1], 'ตำแหน่งNode1' : ListNodeType[0], 'ตำแหน่งNode2' : ListNodeType[1],'ทิศทางการวาง' : 'Y','TypeList' : TypeList})
                        #print(addRowLoadSlab)
                        allBeamDf =  pd.concat([allBeamDf, addRowBeam.to_frame().T], ignore_index=True)  
                        #print(List)
                        List = [Node_df.loc[Rw,'Node No. (int)']]
                        ListStatus = [Node_df.loc[Rw,'สถานะจุดต่อ']]
                        Listpos = [Node_df.loc[Rw,'พิกัด Y (m)  (.2float)']]
                        ListNodeType = ['ใน']
    save_excel_sheet(allBeamDf,filepath,sheetname='allBeamDf', index=False)

if __name__ == "__main__":
    Allbeamdf(filepath = r'C:\ProjectVenv\.venv\ProjectHouse\Exhouse.xlsx')   