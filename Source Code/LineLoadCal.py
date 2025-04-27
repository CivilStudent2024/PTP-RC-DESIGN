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

def allLineLoadcal(filepath,DL,LL) :
    filepath = filepath
    Lineload_df = pd.read_excel(filepath,sheet_name='Lineload_Data')
    Node_df = pd.read_excel(filepath,sheet_name='Node_Data')
    allLineLoadDf = pd.DataFrame({'NodeList1' : [],'NodeList2' : [],'Load(T/M)' : [],'direction' : []})
    for Row,Value in enumerate(Lineload_df['direction']) :
        node1 = int(Lineload_df.loc[Row,'NodeList1'])
        node2 = int(Lineload_df.loc[Row,'NodeList2'])
        Load = float(Lineload_df.loc[Row,'Load(T/M)']) * DL

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
                        allLineLoadDf =  pd.concat([allLineLoadDf, addRowLoadSlab.to_frame().T], ignore_index=True) 
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
                        allLineLoadDf =  pd.concat([allLineLoadDf, addRowLoadSlab.to_frame().T], ignore_index=True) 
                        
                        List = [Node_df.loc[Rw,'Node No. (int)']]
    save_excel_sheet(allLineLoadDf, filepath, sheetname = 'allLineLoadDistributed' , index=False)

if __name__ == "__main__":
    allLineLoadcal(filepath= r'C:\ProjectVenv\.venv\ProjectHouse\Exhouse.xlsx')


