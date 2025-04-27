import pandas as pd 
import os
#from BeamAstCaler import BeamAstcal
import math
from SlabAsCaler import thicknessCheck , Slab_getAs,get_UseST

def save_excel_sheet(df, filepath, sheetname, index=False):
    # Create file if it does not exist
    if not os.path.exists(filepath):
        df.to_excel(filepath, sheet_name=sheetname, index=index)

    # Otherwise, add a sheet. Overwrite if there exists one with the same name.
    else:
        with pd.ExcelWriter(filepath, engine='openpyxl', if_sheet_exists='replace', mode='a') as writer:
            df.to_excel(writer, sheet_name=sheetname, index=index)
def TwoWay_analysed(filepath,DL,LL,fc,fy) :
    Top_Cover = 0.02
    Bot_Cover = 0.02
    Slab_Data_df = pd.read_excel(filepath,sheet_name='Slab_Data')
    Node_df = pd.read_excel(filepath,sheet_name='Node_Data')
    Beam_df = pd.read_excel(filepath,sheet_name='Beam_Data')
    LoadCombine = pd.DataFrame({'DL': [DL],'LL' : [LL]})
    TwoWay_df = pd.DataFrame({'หมายเลขพื้น':[],'NodeI' : [] , 'NodeJ' : [] , 'NodeK' : [] , 'NodeL' : [] , 'Xlenght' : [] , 'Ylenght' : [] , 'm' : [],'Ix' : [],'Iy' : [],'Jx':[],'Jy':[],'Kx' : [],'Ky' : [],'Lx':[],'Ly':[],'DL' :[],'LL':[],'Type' : [] })
    TwoWay_Result_df = pd.DataFrame({})
    #display(Slab_Data_df)


    Interval_df = pd.DataFrame({'Moment' : ['M-Con','M-Discon','M+Mid'], 1.0 : [0.033,0,0.025], 0.9 : [0.040,0,0.030] , 0.8 : [0.055,0,0.041],
                                0.7 : [0.063,0,0.047] , 0.6 : [0.063,0,0.047] , 0.5 : [0.083,0,0.062] , 'ช่วงยาว' : [0.033,0,0.025] })
    OneDisCon_df = pd.DataFrame({'Moment' : ['M-Con','M-Discon','M+Mid'],1.0 : [0.041,0.021,0.031],0.9 : [0.048,0.024,0.036] , 0.8 : [0.055,0.027,0.041],
                                0.7 : [0.062,0.031,0.047] , 0.6 : [0.069,0.035,0.052] , 0.5 : [0.085,0.042,0.064] , 'ช่วงยาว' : [0.041,0.021,0.031] })
    TwoDisCon_df = pd.DataFrame({'Moment' : ['M-Con','M-Discon','M+Mid'],1.0 : [0.049,0.025,0.037],0.9 : [0.057,0.028,0.043] , 0.8 : [0.064,0.032,0.048],
                                0.7 : [0.071,0.036,0.054] , 0.6 : [0.078,0.039,0.059] , 0.5 : [0.090,0.045,0.068] , 'ช่วงยาว' : [0.049,0.025,0.037] })
    ThreeDisCon_df = pd.DataFrame({'Moment' : ['M-Con','M-Discon','M+Mid'],1.0 : [0.058,0.029,0.044],0.9 : [0.066,0.033,0.050] , 0.8 : [0.074,0.037,0.056],
                                0.7 : [0.082,0.041,0.062] , 0.6 : [0.090,0.045,0.068] , 0.5 : [0.098,0.049,0.074] , 'ช่วงยาว' : [0.058,0.029,0.044] })
    FourDiscon_df = pd.DataFrame({'Moment' : ['M-Con','M-Discon','M+Mid'],1.0 : [0,0.033,0.050],0.9 : [0,0.038,0.057] , 0.8 : [0,0.043,0.064],
                                0.7 : [0,0.047,0.072] , 0.6 : [0,0.053,0.080] , 0.5 : [0,0.055,0.083] , 'ช่วงยาว' : [0,0.033,0.050] })


    for index, slabtype in enumerate(Slab_Data_df['ประเภทแผ่นพื้น']) :
        if slabtype == 'พื้นหล่อในที่' :
            NodeI = Slab_Data_df.loc[index,'จุดต่อ I'] ; NodeJ = Slab_Data_df.loc[index,'จุดต่อ J'] ;  NodeK = Slab_Data_df.loc[index,'จุดต่อ K'] ;  NodeL = Slab_Data_df.loc[index,'จุดต่อ L']
            Xlenght = float(Slab_Data_df.loc[index,'Jx']) - float(Slab_Data_df.loc[index,'Ix'])
            Ylenght = float(Slab_Data_df.loc[index,'Ky']) - float(Slab_Data_df.loc[index,'Jy'])
            no_slab = int(Slab_Data_df.loc[index,'หมายเลขพื้น'])
            m = min([Xlenght,Ylenght]) / max([Xlenght,Ylenght])
            DL = float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกคงที่อื่นๆ'])
            LL = float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกจรอื่นๆ'])
            for index,Node_i in enumerate(Node_df['Node No. (int)']) :
                if Node_i == NodeI :
                    Ix = Node_df.loc[index,'พิกัด X (m)(.2float)']
                    Iy = Node_df.loc[index,'พิกัด Y (m)  (.2float)']
                elif Node_i == NodeJ :
                    Jx = Node_df.loc[index,'พิกัด X (m)(.2float)']
                    Jy = Node_df.loc[index,'พิกัด Y (m)  (.2float)']
                elif Node_i == NodeK :
                    Kx = Node_df.loc[index,'พิกัด X (m)(.2float)']
                    Ky = Node_df.loc[index,'พิกัด Y (m)  (.2float)']
                elif Node_i == NodeL :
                    Lx = Node_df.loc[index,'พิกัด X (m)(.2float)']
                    Ly = Node_df.loc[index,'พิกัด Y (m)  (.2float)']


            if 0.5 <= m <= 1.0 : #พื้นสองทาง

                new_row = pd.Series({'หมายเลขพื้น':no_slab,'NodeI' : NodeI , 'NodeJ' : NodeJ , 'NodeK' : NodeK , 'NodeL' : NodeL , 'Xlenght' : Xlenght , 'Ylenght' : Ylenght , 'm' : m,'Ix' : Ix,'Iy' : Iy,'Jx':Jx,'Jy':Jy,'Kx' : Kx,'Ky' : Ky,'Lx':Lx,'Ly':Ly,'DL':DL,'LL' : LL,'Type' : 'Two-Way' })
                TwoWay_df = pd.concat([TwoWay_df, new_row.to_frame().T], ignore_index=True)
            else :
                new_row = pd.Series({'หมายเลขพื้น':no_slab,'NodeI' : NodeI , 'NodeJ' : NodeJ , 'NodeK' : NodeK , 'NodeL' : NodeL , 'Xlenght' : Xlenght , 'Ylenght' : Ylenght , 'm' : m,'Ix' : Ix,'Iy' : Iy,'Jx':Jx,'Jy':Jy,'Kx' : Kx,'Ky' : Ky,'Lx':Lx,'Ly':Ly,'DL':DL,'LL' : LL,'Type' : 'One-Way' })
                TwoWay_df = pd.concat([TwoWay_df, new_row.to_frame().T], ignore_index=True)

    for index, row in TwoWay_df.iterrows():
        if row['Type'] == 'Two-Way' :
            W = [int(row['NodeI']),int(row['NodeL'])] ; S = [int(row['NodeI']), int(row['NodeJ'])] ; E = [int(row['NodeJ']),int(row['NodeK'])] ; N = [int(row['NodeL']) , int(row['NodeK'])]
            Xlenght = float(row['Xlenght']) ; Ylenght = float(row['Ylenght'])
            DL = row['DL'] ; LL =row['LL']
            Case = '4-Disconnected'
            ConnectionList = []
            check = 0

            for i,rowx in Node_df.iterrows() :
                if rowx['Node No. (int)'] == W[0] :
                    xw = float(rowx['พิกัด X (m)(.2float)'])
                    yw1 = float(rowx['พิกัด Y (m)  (.2float)'])
                elif  rowx['Node No. (int)'] == W[1] :
                    yw2 = float(rowx['พิกัด Y (m)  (.2float)']) 
                if rowx['Node No. (int)'] == S[0] :
                    ys = float(rowx['พิกัด Y (m)  (.2float)'])
                    xs1 = float(rowx['พิกัด X (m)(.2float)'])
                elif  rowx['Node No. (int)'] == S[1] :
                    xs2 = float(rowx['พิกัด X (m)(.2float)']) 
                if rowx['Node No. (int)'] == E[0] :
                    xe = float(rowx['พิกัด X (m)(.2float)'])
                    ye1 = float(rowx['พิกัด Y (m)  (.2float)'])
                elif  rowx['Node No. (int)'] == E[1] :
                    ye2 = float(rowx['พิกัด Y (m)  (.2float)']) 
                if rowx['Node No. (int)'] == N[0] :
                    yn = float(rowx['พิกัด Y (m)  (.2float)'])
                    xn1 = float(rowx['พิกัด X (m)(.2float)'])
                elif  rowx['Node No. (int)'] == N[1] :
                    xn2 = float(rowx['พิกัด X (m)(.2float)']) 

            #W
            for node,rw in TwoWay_df.iterrows() :
                if rw['Jx'] == xw and rw['Kx'] == xw  and (rw['Jy'] <= yw1 <= rw['Ky']) and (rw['Jy'] <= yw2 <= rw['Ky']) and  (rw['หมายเลขพื้น'] != row['หมายเลขพื้น']): #W คือ x ค่าตรง y เปลี่ยน
                    if rw['Ylenght'] > 0.7*Ylenght :
                        check = 1            
            if check == 1 :            
                ConnectionList.append('C') ; W.append('C')   
            else : ConnectionList.append('D') ; W.append('D') 
            check = 0 
            for node,rw in TwoWay_df.iterrows() :    
                if rw['Ly'] == ys and rw['Ky'] == ys  and (rw['Lx'] <= xs1 <= rw['Kx']) and (rw['Lx'] <= xs2 <= rw['Kx'])  and  (rw['หมายเลขพื้น'] != row['หมายเลขพื้น']) :
                    if rw['Xlenght'] > 0.7*Xlenght : #0.7 คือความยาวที่พิจารณาว่าเป็นคานต่อเนื่อง
                        check = 1
            if check == 1 :
                ConnectionList.append('C') ; S.append('C')     
            else : ConnectionList.append('D') ; S.append('D') 
            check = 0
            for node,rw in TwoWay_df.iterrows() : 
                if rw['Ix'] == xe and rw['Lx'] == xe  and (rw['Iy'] <= ye1 <= rw['Ly']) and (rw['Iy'] <= ye2 <= rw['Ly']) and  (rw['หมายเลขพื้น'] != row['หมายเลขพื้น']) :
                    if rw['Ylenght'] > 0.7*Ylenght : 
                        check = 1
            if check == 1 :
                ConnectionList.append('C') ; E.append('C') 
            else : ConnectionList.append('D')  ; E.append('D') 
            check = 0

            for node,rw in TwoWay_df.iterrows() : 
                if rw['Iy'] == yn and rw['Jy'] == yn  and (rw['Ix'] <= xn1 <= rw['Jx']) and (rw['Ix'] <= xn2 <= rw['Jx']) and  (rw['หมายเลขพื้น'] != row['หมายเลขพื้น'])  :
                    if rw['Xlenght'] > 0.7*Xlenght : 
                        check = 1
            if check == 1 :
                ConnectionList.append('C') ; N.append('C')  
            else : ConnectionList.append('D') ; N.append('D')  
            check = 0 

            #print(ConnectionList)
            #print(W,N,E,S)
            #print(W,N,E,S)
            match ConnectionList.count('C') :
                case 0 : C_df = FourDiscon_df
                case 1 : C_df = ThreeDisCon_df ; Case = '3-Disconnected'
                case 2 : C_df = TwoDisCon_df ; Case = '2-Disconnected'
                case 3 : C_df = OneDisCon_df ; Case = '1-Disconnected'
                case 4 : C_df = Interval_df ; Case = 'Interval'
            #display(C_df)

            Coeff_list = {'S' : [] ,'V' : [] ,'N' : [] ,'W' : [] ,'H' : [] , 'E':[] }
            for i in [S,N,W,E] : #S,V,N,W,H,E  ตามลำดับ
                m = float(TwoWay_df.loc[index,'m'])
                if i == W or i == E : # Ylenght
                    if i[2] == 'C':
                        if Xlenght <= Ylenght : #ด้าน W,E เป็นช่วงสั้น
                            C = C_df.loc[0,math.floor(m*10**1)/10**1]
                            if i == W :
                                Coeff_list['W'] = C
                                W.append('Short')
                            else : Coeff_list['E'] = C ; E.append('Short')

                        else: 
                            C = C_df.loc[0,'ช่วงยาว']
                            if i == W :
                                Coeff_list['W'] = C
                                W.append('Long')
                            else : Coeff_list['E'] = C ; E.append('Long')
                    else : 
                        if Xlenght <= Ylenght :
                            C = C_df.loc[1,math.floor(m*10**1)/10**1]
                            if i == W :
                                Coeff_list['W'] = C
                                W.append('Short')
                            else : Coeff_list['E'] = C ; E.append('Short')                    
                        else :
                            C = C_df.loc[1,'ช่วงยาว']
                            if i == W :
                                Coeff_list['W'] = C
                                W.append('Long')
                            else : Coeff_list['E'] = C ; E.append('Long')   
            


                elif i == N or i == S : # Xlenght
                    if i[2] == 'C' :
                        if Ylenght <= Xlenght :
                            C = C_df.loc[0,math.floor(m*10**1)/10**1]

                            if i == N :
                                Coeff_list['N'] = C
                                N.append('Short')
                            else : Coeff_list['S'] = C ; S.append('Short') 


                        else:
                            C = C_df.loc[0,'ช่วงยาว']
                            if i == N :
                                Coeff_list['N'] = C
                                N.append('Long')
                            else : Coeff_list['S'] = C ; S.append('Long')    

                    else : 
                        if Ylenght <= Xlenght :
                            C = C_df.loc[1,math.floor(m*10**1)/10**1]
                            if i == N :
                                Coeff_list['N'] = C
                                N.append('Short')
                            else : Coeff_list['S'] = C ; S.append('Short')    
                        else :
                            C = C_df.loc[1,'ช่วงยาว']
                            if i == N :
                                Coeff_list['N'] = C
                                N.append('Long')
                            else : Coeff_list['S'] = C ; S.append('Long')
            Side_V = None
            Side_H = None
            if Ylenght <= Xlenght : #WE ช่วงสั้น
                Coeff_list['H'] = C_df.loc[2,math.floor(m*10**1)/10**1]
                Coeff_list['V'] = C_df.loc[2,'ช่วงยาว']
                Side_H = 'Short' ; Side_V = 'Long'
            else :
                Coeff_list['V'] = C_df.loc[2,math.floor(m*10**1)/10**1]
                Coeff_list['H'] = C_df.loc[2,'ช่วงยาว']
                Side_H = 'Long' ; Side_V = 'Short'
            print(W,E,N,S)

            Moment_list = {'S' : 0 ,'V' : 0 ,'N' : 0 ,'W' : 0 ,'H' : 0 , 'E': 0 }
            thickness_temp = math.ceil((((2*Xlenght)+(2*Ylenght))/180)*100)/100  # ความหนาเริ่มต้น
            if thickness_temp < 0.1 :
                thickness_temp = 0.1 
            for check in range(0,5) : # ตรวจสอบความถูกต้องด้วยการวนลูป
                for i in Moment_list :
                    thickness_initial = thickness_temp
                    DL = 2.4*thickness_initial
                    Moment_list[i]  = Coeff_list[i] * ( (DL*(float(LoadCombine.loc[0,'DL']))) + (LL*(float(LoadCombine.loc[0,'LL']))) ) * ( min([Xlenght,Ylenght])  **2 )      
                    Mmax = max(Moment_list.values())
                    Vmax = ((DL*(float(LoadCombine.loc[0,'DL']))) + (LL*(float(LoadCombine.loc[0,'LL']))) * min([Xlenght,Ylenght])) / 4
                    thickness_last = thicknessCheck(Mu = Mmax,Vu = Vmax,fc = fc,fy = fy,botcover = 0.02,thickness=thickness_initial,safetyfactor=0.9)
                    thickness_temp = thickness_last
                if thickness_initial == thickness_last :
                    break

            thickness_use = round(thickness_last,2)
            DL = 2.4*thickness_use
            for i in Moment_list :
                Moment_list[i]  = float(Coeff_list[i] * ( (DL*(float(LoadCombine.loc[0,'DL']))) + (LL*(float(LoadCombine.loc[0,'LL']))) ) * ( min([Xlenght,Ylenght])  **2 ))

            
            As_list = {'S' : 0 ,'V' : 0 ,'N' : 0 ,'W' : 0 ,'H' : 0 , 'E': 0 }
            side_list = {'S' : S[3] , 'V' : Side_V ,  'N' : N[3],'W' : W[3],'H' : Side_H,'E' : E[3]}
            for i in As_list :
                As_list[i] = round(Slab_getAs(Mu = Moment_list[i],fc = fc,fy = fy,botcover = 0.02,thickness=thickness_use,safetyfactor=0.9,side=side_list[i]),2)
                
            
            print('Coeff is ',Coeff_list) 
            print('Moment is ',Moment_list)
            print('As is',As_list)
            
            newrow_result = pd.Series({'หมายเลขพื้น' : row['หมายเลขพื้น'],'SlabType' : 'Two-Way','Case' : Case ,'Xlength' : Xlenght,'Ylength' : Ylenght, 'Thickness' : f'{thickness_use:.2f}','Top Cover' : Top_Cover ,'Bot Cover' : Bot_Cover ,'AS#1':As_list['S'],'AS#2':As_list['V'],'AS#3':As_list['N'],'AS#4':As_list['W'],'AS#5':As_list['H'],
                                    'AS#6':As_list['E'],'Use ST#1' : get_UseST(As=As_list['S'],thickness=thickness_use,Mu=Moment_list['S'],fc=fc,side = side_list['S']),'Use ST#2' : get_UseST(As=As_list['V'],thickness=thickness_use,Mu=Moment_list['V'],fc=fc,side = Side_V),
                                    'Use ST#3' : get_UseST(As=As_list['N'],thickness=thickness_use,Mu=Moment_list['N'],fc=fc,side = side_list['N']),'Use ST#4' : get_UseST(As=As_list['W'],thickness=thickness_use,Mu=Moment_list['W'],fc=fc,side = side_list['W']),'Use ST#5' : get_UseST(As=As_list['H'],thickness=thickness_use,Mu=Moment_list['H'],side = Side_H,fc=fc),'Use ST#6' : get_UseST(As=As_list['E'],thickness=thickness_use,Mu=Moment_list['E'],fc=fc,side = side_list['E'])})

            TwoWay_Result_df = pd.concat([TwoWay_Result_df, newrow_result.to_frame().T], ignore_index=True)
    return TwoWay_Result_df
def OneWay_analysed(filepath,DL,LL,fc,fy) :
    fc=fc
    fy=fy
    Top_Cover = 0.02
    Bot_Cover = 0.02
    Slab_Data_df = pd.read_excel(filepath,sheet_name='Slab_Data')
    Node_df = pd.read_excel(filepath,sheet_name='Node_Data')
    Beam_df = pd.read_excel(filepath,sheet_name='Beam_Data')
    Type_df = pd.read_excel(filepath,sheet_name='SectionType')
    LoadCombine = pd.DataFrame({'DL': [DL],'LL' : [LL]})
    OneWay_df = pd.DataFrame({'หมายเลขพื้น':[],'NodeI' : [] , 'NodeJ' : [] , 'NodeK' : [] , 'NodeL' : [] , 'Xlenght' : [] , 'Ylenght' : [] , 'm' : [],'Ix' : [],'Iy' : [],'Jx':[],'Jy':[],'Kx' : [],'Ky' : [],'Lx':[],'Ly':[],'DL' :[],'LL':[],'Type' : [] })
    OneWay_Result_df = pd.DataFrame({})
    #display(Slab_Data_df)


    Interval_df = {'M-Con' : 0.0909091 ,'M-Discon' : 0,'M-Mid' :  0.0625 }
    OneDisCon_df =  {'M-Con' : 0.111111,'M-Discon' : 0.04167 ,'M-Mid' :  0.07143 }
    TwoDisCon_df =  {'M-Con' : 0 ,'M-Discon' : 0 ,'M-Mid' :  0.125 }



    for index, slabtype in enumerate(Slab_Data_df['ประเภทแผ่นพื้น']) :
        NodeI = Slab_Data_df.loc[index,'จุดต่อ I'] ; NodeJ = Slab_Data_df.loc[index,'จุดต่อ J'] ;  NodeK = Slab_Data_df.loc[index,'จุดต่อ K'] ;  NodeL = Slab_Data_df.loc[index,'จุดต่อ L']
        Xlenght = float(Slab_Data_df.loc[index,'Jx']) - float(Slab_Data_df.loc[index,'Ix'])
        Ylenght = float(Slab_Data_df.loc[index,'Ky']) - float(Slab_Data_df.loc[index,'Jy'])
        no_slab = int(Slab_Data_df.loc[index,'หมายเลขพื้น'])
        if slabtype == 'พื้นหล่อในที่' :
            m = min([Xlenght,Ylenght]) / max([Xlenght,Ylenght])
            DL = float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกคงที่อื่นๆ'])
            LL = float(Slab_Data_df.loc[index,'น้ำหนักบรรทุกจรอื่นๆ'])
            for index,Node_i in enumerate(Node_df['Node No. (int)']) :
                if Node_i == NodeI :
                    Ix = Node_df.loc[index,'พิกัด X (m)(.2float)']
                    Iy = Node_df.loc[index,'พิกัด Y (m)  (.2float)']
                elif Node_i == NodeJ :
                    Jx = Node_df.loc[index,'พิกัด X (m)(.2float)']
                    Jy = Node_df.loc[index,'พิกัด Y (m)  (.2float)']
                elif Node_i == NodeK :
                    Kx = Node_df.loc[index,'พิกัด X (m)(.2float)']
                    Ky = Node_df.loc[index,'พิกัด Y (m)  (.2float)']
                elif Node_i == NodeL :
                    Lx = Node_df.loc[index,'พิกัด X (m)(.2float)']
                    Ly = Node_df.loc[index,'พิกัด Y (m)  (.2float)']


            if m < 0.5 or m > 2 : #พื้นทางเดียว
                new_row = pd.Series({'หมายเลขพื้น':no_slab,'NodeI' : NodeI , 'NodeJ' : NodeJ , 'NodeK' : NodeK , 'NodeL' : NodeL , 'Xlenght' : Xlenght , 'Ylenght' : Ylenght , 'm' : m,'Ix' : Ix,'Iy' : Iy,'Jx':Jx,'Jy':Jy,'Kx' : Kx,'Ky' : Ky,'Lx':Lx,'Ly':Ly,'DL':DL,'LL' : LL,'Type' : 'One-Way' })
                OneWay_df = pd.concat([OneWay_df, new_row.to_frame().T], ignore_index=True)
            else :
                new_row = pd.Series({'หมายเลขพื้น':no_slab,'NodeI' : NodeI , 'NodeJ' : NodeJ , 'NodeK' : NodeK , 'NodeL' : NodeL , 'Xlenght' : Xlenght , 'Ylenght' : Ylenght , 'm' : m,'Ix' : Ix,'Iy' : Iy,'Jx':Jx,'Jy':Jy,'Kx' : Kx,'Ky' : Ky,'Lx':Lx,'Ly':Ly,'DL':DL,'LL' : LL,'Type' : 'Two-Way' })
                OneWay_df = pd.concat([OneWay_df, new_row.to_frame().T], ignore_index=True)                
        elif slabtype == 'แผ่นพื้นสำเร็จวางแนวนอน' :
            newrow_result = pd.Series({'หมายเลขพื้น' : no_slab,'SlabType' : 'One-Way','Case' : 'X-Direction' ,'Xlength' : Xlenght,'Ylength' : Ylenght, 'Thickness' : '' ,'AS#1':'','AS#2':'','AS#3':'','AS#4':'','AS#5':'',
                                    'AS#6': '','Use ST#1' : '','Use ST#2' : '' ,
                                    'Use ST#3' : '','Use ST#4' : '','Use ST#5' :'','Use ST#6' : ''})
            OneWay_Result_df = pd.concat([OneWay_Result_df, newrow_result.to_frame().T], ignore_index=True)            

        else :
            newrow_result = pd.Series({'หมายเลขพื้น' : no_slab,'SlabType' : 'One-Way','Case' : 'Y-Direction', 'Xlength' : Xlenght,'Ylength' : Ylenght, 'Thickness' : '' ,'AS#1':'','AS#2':'','AS#3':'','AS#4':'','AS#5':'',
                                    'AS#6': '','Use ST#1' : '','Use ST#2' : '' ,
                                    'Use ST#3' : '','Use ST#4' : '','Use ST#5' :'' ,'Use ST#6' : ''})
            OneWay_Result_df = pd.concat([OneWay_Result_df, newrow_result.to_frame().T], ignore_index=True)  
    #display(OneWay_df)
    for index, row in OneWay_df.iterrows():
        if row['Type'] == 'One-Way' :
            W = [int(row['NodeI']),int(row['NodeL'])] ; S = [int(row['NodeI']), int(row['NodeJ'])] ; E = [int(row['NodeJ']),int(row['NodeK'])] ; N = [int(row['NodeL']) , int(row['NodeK'])]

            Xsize_s = 0
            Xsize_n = 0
            Xsize_w = 0
            Xsize_e = 0
            node1x = int(row['Ix'])
            node2x = int(row['Jx'])
            for indexbeam_find , rowbeam_find in Beam_df.iterrows() :
                x1 = float(rowbeam_find['Node1x']) ; x2 = float(rowbeam_find['Node2x'])
                if x1 <= node1x <=x2 and x1 <= node2x <= x2 :
                    TypeList_s = int(rowbeam_find['หน้าตัดคาน'])
                    for it,t in enumerate(Type_df['หมายเลขหน้าตัดคาน']) :
                        if TypeList_s == t :
                            Xsize_s = float(Type_df.loc[it,'ความกว้าง (m)'])
                            break
                    break
            node1x = int(row['Lx'])
            node2x = int(row['Kx'])
            for indexbeam_find , rowbeam_find in Beam_df.iterrows() :
                x1 = float(rowbeam_find['Node1x']) ; x2 = float(rowbeam_find['Node2x'])
                if x1 <= node1x <=x2 and x1 <= node2x <= x2 :
                    TypeList_s = int(rowbeam_find['หน้าตัดคาน'])
                    for it,t in enumerate(Type_df['หมายเลขหน้าตัดคาน']) :
                        if TypeList_s == t :
                            Xsize_n = float(Type_df.loc[it,'ความกว้าง (m)'])
                            break
                    break

            node1x = int(row['Iy'])
            node2x = int(row['Ly'])
            for indexbeam_find , rowbeam_find in Beam_df.iterrows() :
                x1 = float(rowbeam_find['Node1y']) ; x2 = float(rowbeam_find['Node2y'])
                if x1 <= node1x <=x2 and x1 <= node2x <= x2 :
                    TypeList_s = int(rowbeam_find['หน้าตัดคาน'])
                    for it,t in enumerate(Type_df['หมายเลขหน้าตัดคาน']) :
                        if TypeList_s == t :
                            Xsize_w = float(Type_df.loc[it,'ความกว้าง (m)'])
                            break
                    break
            node1x = int(row['Jy'])
            node2x = int(row['Ky'])
            for indexbeam_find , rowbeam_find in Beam_df.iterrows() :
                x1 = float(rowbeam_find['Node1y']) ; x2 = float(rowbeam_find['Node2y'])
                if x1 <= node1x <=x2 and x1 <= node2x <= x2 :
                    TypeList_s = int(rowbeam_find['หน้าตัดคาน'])
                    for it,t in enumerate(Type_df['หมายเลขหน้าตัดคาน']) :
                        if TypeList_s == t :
                            Xsize_e = float(Type_df.loc[it,'ความกว้าง (m)'])
                            break
                    break

            Xlenght = float(row['Xlenght']) ; Ylenght = float(row['Ylenght'])
            Xeff = Xlenght - Xsize_w - Xsize_e ; Yeff = Ylenght - Xsize_n - Xsize_s
            DL = row['DL'] ; LL =row['LL']
            Case = '2-Disconnected'
            ConnectionList = []
            check = 0

            for i,rowx in Node_df.iterrows() :
                if rowx['Node No. (int)'] == W[0] :
                    xw = float(rowx['พิกัด X (m)(.2float)'])
                    yw1 = float(rowx['พิกัด Y (m)  (.2float)'])
                elif  rowx['Node No. (int)'] == W[1] :
                    yw2 = float(rowx['พิกัด Y (m)  (.2float)']) 
                if rowx['Node No. (int)'] == S[0] :
                    ys = float(rowx['พิกัด Y (m)  (.2float)'])
                    xs1 = float(rowx['พิกัด X (m)(.2float)'])
                elif  rowx['Node No. (int)'] == S[1] :
                    xs2 = float(rowx['พิกัด X (m)(.2float)']) 
                if rowx['Node No. (int)'] == E[0] :
                    xe = float(rowx['พิกัด X (m)(.2float)'])
                    ye1 = float(rowx['พิกัด Y (m)  (.2float)'])
                elif  rowx['Node No. (int)'] == E[1] :
                    ye2 = float(rowx['พิกัด Y (m)  (.2float)']) 
                if rowx['Node No. (int)'] == N[0] :
                    yn = float(rowx['พิกัด Y (m)  (.2float)'])
                    xn1 = float(rowx['พิกัด X (m)(.2float)'])
                elif  rowx['Node No. (int)'] == N[1] :
                    xn2 = float(rowx['พิกัด X (m)(.2float)']) 

            #W
            for node,rw in OneWay_df.iterrows() :
                if rw['Jx'] == xw and rw['Kx'] == xw  and (rw['Jy'] <= yw1 <= rw['Ky']) and (rw['Jy'] <= yw2 <= rw['Ky']) and  (rw['หมายเลขพื้น'] != row['หมายเลขพื้น']): #W คือ x ค่าตรง y เปลี่ยน
                    if float(rw['Xlenght']) > 0.7*Xlenght and ((yw2-yw1) == max([Xlenght,Ylenght])) :
                        check = 1            
            if check == 1 :            
                ConnectionList.append('C') ; W.append('C')   
            else : ConnectionList.append('D') ; W.append('D') 
            check = 0 
            for node,rw in OneWay_df.iterrows() :    
                if rw['Ly'] == ys and rw['Ky'] == ys  and (rw['Lx'] <= xs1 <= rw['Kx']) and (rw['Lx'] <= xs2 <= rw['Kx'])  and  (rw['หมายเลขพื้น'] != row['หมายเลขพื้น']) :
                    if rw['Ylenght'] > 0.7*Ylenght and ((xs2-xs1) == max([Xlenght,Ylenght])): 
                        check = 1
            if check == 1 :
                ConnectionList.append('C') ; S.append('C')     
            else : ConnectionList.append('D') ; S.append('D') 
            check = 0
            for node,rw in OneWay_df.iterrows() : 
                if rw['Ix'] == xe and rw['Lx'] == xe  and (rw['Iy'] <= ye1 <= rw['Ly']) and (rw['Iy'] <= ye2 <= rw['Ly']) and  (rw['หมายเลขพื้น'] != row['หมายเลขพื้น']) :
                    if rw['Xlenght'] > 0.7*Xlenght and ((ye2-ye1) == max([Xlenght,Ylenght])) : 
                        check = 1
            if check == 1 :
                ConnectionList.append('C') ; E.append('C') 
            else : ConnectionList.append('D')  ; E.append('D') 
            check = 0

            for node,rw in OneWay_df.iterrows() : 
                if rw['Iy'] == ys and rw['Jy'] == ys  and (rw['Ix'] <= xn1 <= rw['Jx']) and (rw['Ix'] <= xn2 <= rw['Jx']) and  (rw['หมายเลขพื้น'] != row['หมายเลขพื้น'])  :
                    if rw['Ylenght'] > 0.7*Ylenght and ((xn2-xn1) == max([Xlenght,Ylenght])) : 
                        check = 1
            if check == 1 :
                ConnectionList.append('C') ; N.append('C')  
            else : ConnectionList.append('D') ; N.append('D')  
            check = 0 

            #print(ConnectionList)
            #print(W,N,E,S)
            #print(W,N,E,S)
            match ConnectionList.count('C') :

                case 0 : C_df = TwoDisCon_df ; Case = '2-Disconnected'
                case 1 : C_df = OneDisCon_df ; Case = '1-Disconnected'
                case 2 : C_df = Interval_df ; Case = 'Interval'

            #display(C_df)

            Coeff_list = {'S' : 0 ,'V' : 0 ,'N' : 0 ,'W' : 0 ,'H' : 0 , 'E':0 }

            C_con = C_df['M-Con']
            C_discon = C_df['M-Discon']
            if Ylenght <= Xlenght : #WE ช่วงสั้น
                if N[2] == 'C' :
                    Coeff_list['N'] = C_con
                else :  Coeff_list['N'] = C_discon   
                if S[2] == 'C' :
                    Coeff_list['S'] = C_con
                else :  Coeff_list['S'] = C_discon   

                Coeff_list['V'] = C_df['M-Mid']
                Coeff_list['H'] = 0
            else :

                if W[2] == 'C' :
                    Coeff_list['W'] = C_con
                else :  Coeff_list['W'] = C_discon   
                if E[2] == 'C' :
                    Coeff_list['E'] = C_con
                else :  Coeff_list['E'] = C_discon  
                Coeff_list['H'] = C_df['M-Mid']
                Coeff_list['V'] = 0

            Moment_list = {'S' : 0 ,'V' : 0 ,'N' : 0 ,'W' : 0 ,'H' : 0 , 'E':0 }
            print(Coeff_list)
            if Case == '2-Disconnected' :
                thickness_temp = math.floor(min([Xlenght,Ylenght]) / 20)  
            elif Case == '1-Disconnected' :
                thickness_temp = math.floor(min([Xlenght,Ylenght]) / 24)
            else : thickness_temp = math.floor(min([Xlenght,Ylenght]) / 28)
            if thickness_temp < 0.1 :
                thickness_temp =0.1
            for check in range(0,5) : # ตรวจสอบความถูกต้องด้วยการวนลูป
                for i in Moment_list :
                    thickness_initial = thickness_temp
                    DeadLoad = DL + 2.4*thickness_initial
                    print('DeadLoad is',DeadLoad,'Live Load is',LL)
                    C = Coeff_list[i]
                    Moment_list[i]  = float(C) * ( (DeadLoad*(float(LoadCombine.loc[0,'DL']))) + (LL*(float(LoadCombine.loc[0,'LL']))) ) * ( min([Xeff,Yeff])  **2 )      
                    Mmax = max(Moment_list.values())
                    Vmax = ((DeadLoad*(float(LoadCombine.loc[0,'DL']))) + (LL*(float(LoadCombine.loc[0,'LL']))) * min([Xlenght,Ylenght])) * 1.15 / 2
                    thickness_last = thicknessCheck(Mu = Mmax,Vu = Vmax,fc = fc,fy = fy,botcover = 0.02,thickness=thickness_initial,safetyfactor=0.9)
                    thickness_temp = thickness_last
                if thickness_initial == thickness_last :
                    break

            thickness_use = round(thickness_last,2)



            As_list = {'S' : 0 ,'V' : 0 ,'N' : 0 ,'W' : 0 ,'H' : 0 , 'E':0 }
            for i in As_list :
                As_list[i] = round(Slab_getAs(Mu = Moment_list[i],fc = fc,fy = fy,botcover = 0.02,thickness=thickness_use,safetyfactor=0.9),2)

            print(N,W,E,S)
            print(Coeff_list) 
            print(Moment_list)
            print(As_list)
            if Ylenght <= Xlenght :
                newrow_result = pd.Series({'หมายเลขพื้น' : row['หมายเลขพื้น'],'SlabType' : 'One-Way','Case' : Case ,'Xlength' : Xlenght,'Ylength' : Ylenght, 'Thickness' : f'{thickness_use:.2f}','Top Cover' : Top_Cover ,'Bot Cover' : Bot_Cover ,'AS#1':As_list['S'],'AS#2':As_list['V'],'AS#3':As_list['N'],'AS#4':As_list['W'],'AS#5':As_list['H'],
                                        'AS#6':As_list['E'],'Use ST#1' : get_UseST(As=As_list['S'],thickness=thickness_use,Mu=Moment_list['S'],fc=fc),'Use ST#2' : get_UseST(As=As_list['V'],thickness=thickness_use,Mu=Moment_list['V'],fc=fc),
                                        'Use ST#3' : get_UseST(As=As_list['N'],thickness=thickness_use,Mu=Moment_list['N'],fc=fc),'Use ST#4' : get_UseST(As=As_list['W'],thickness=thickness_use,Mu=Moment_list['W'],fc=fc),'Use ST#5' : get_UseST(As=As_list['H'],thickness=thickness_use,Mu=Moment_list['H'],fc=fc),'Use ST#6' : get_UseST(As=As_list['E'],thickness=thickness_use,Mu=Moment_list['E'],fc=fc)})
            else :
                newrow_result = pd.Series({'หมายเลขพื้น' : row['หมายเลขพื้น'],'SlabType' : 'One-Way','Case' : Case ,'Xlength' : Xlenght,'Ylength' : Ylenght, 'Thickness' : f'{thickness_use:.2f}','Top Cover' : Top_Cover ,'Bot Cover' : Bot_Cover ,'AS#1':As_list['S'],'AS#2':As_list['V'],'AS#3':As_list['N'],'AS#4': As_list['W'] ,'AS#5':As_list['H'],
                                        'AS#6':As_list['E'],'Use ST#1' : get_UseST(As=As_list['S'],thickness=thickness_use,Mu=Moment_list['S'],fc=fc),'Use ST#2' : get_UseST(As=As_list['V'],thickness=thickness_use,Mu=Moment_list['V'],fc=fc),
                                        'Use ST#3' : get_UseST(As=As_list['N'],thickness=thickness_use,Mu=Moment_list['N'],fc=fc),'Use ST#4' : get_UseST(As=As_list['W'],thickness=thickness_use,Mu=Moment_list['W'],fc=fc),'Use ST#5' : get_UseST(As=As_list['H'],thickness=thickness_use,Mu=Moment_list['H'],fc=fc),'Use ST#6' : get_UseST(As=As_list['E'],thickness=thickness_use,Mu=Moment_list['E'],fc=fc) })
            
            OneWay_Result_df = pd.concat([OneWay_Result_df, newrow_result.to_frame().T], ignore_index=True)
    return OneWay_Result_df

if __name__ == "__main__":
    a = TwoWay_analysed(filepath = r'C:\ProjectVenv\.venv\ProjectHouse\Ex6_3.xlsx',DL = 1.4,LL = 1.7,fc=240,fy=4000)
    b = OneWay_analysed(filepath = r'C:\ProjectVenv\.venv\ProjectHouse\Ex6_3.xlsx',DL = 1.4,LL = 1.7,fc=240,fy=4000)
    all_slab = pd.concat([a,b],ignore_index = True)
    all_slab = all_slab.sort_values(by='หมายเลขพื้น')
    save_excel_sheet(all_slab, filepath = r'C:\ProjectVenv\.venv\ProjectHouse\Ex6_3-Analysed.xlsx', sheetname='Slab_Result', index=False)
    print(all_slab)