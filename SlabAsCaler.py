import math


def thicknessCheck(Mu,Vu,fc,fy,botcover,thickness,safetyfactor) :
    As_List = {'RB6' : 0.28 , 'RB9' : 0.64 ,'DB10' : 0.79}
    depth = thickness - botcover
    beta = 0.85
    Mu = abs(Mu) * 10**5 # T/m to kg-cm
    fc = fc 
    fy = fy #ksc
    pb = (0.85*fc/fy) * beta * (6120/(6120+fy))
    pmax = 0.75*pb
    pmin = 0.0018*100*depth

    Mn = Mu/safetyfactor
    Vn = Vu/0.85
    Ru = Mu / (safetyfactor*100* ((depth*100)**2))

    try :
        p = (0.85*fc/fy) * (1- math.sqrt(1- (2*Ru / (0.85*fc))))
        if p > pmax   :
            firstcheck = 0
            while p > pmax :
                if firstcheck == 0 :
                    depth += 0.01
                    firstcheck = 1
                else :
                    depth += 0.02
                Ru = Mu / (safetyfactor*100* ((depth*100)**2))
                p = (0.85*fc/fy) * (1- math.sqrt(1- (2*Ru / (0.85*fc))))



        else : pass
    except :
        while 1 - (2*Ru / (0.85*fc)) < 0:
            depth += 0.02
            Ru = Mu / (safetyfactor*100* ((depth*100)**2))

    if Vn > 0.53 * math.sqrt(fc) * 100 * depth :
        while Vn > 0.53 * math.sqrt(fc) * 100 * depth :
            depth += 0.02




    return depth + botcover

def Slab_getAs(Mu,fc,fy,botcover,thickness,safetyfactor,side = None) :
    botcover = 0.020
    if side == None or side == 'Short' :
        depth = thickness - botcover - 0.005
    else : depth = thickness - botcover - 0.015

    Mu = abs(Mu) * 10**5 # T/m to kg-cm

    fc = fc 
    fy = fy #ksc
    Mn = Mu/safetyfactor
    Ru = Mu / (safetyfactor*100* ((depth*100)**2))
    p = (0.85*fc/fy) * (1- math.sqrt(1- (2*Ru / (0.85*fc))))
    Asmin = 0.0018 * 100 * thickness*100
    As = p * 100 * depth *100

    if As < Asmin :
        As = Asmin
    return As
def get_UseST(As,thickness,Mu,fc,side= None) :
    if side == None or side == 'Short' :
        depth = thickness  -0.02 - 0.005
    else : depth = thickness  - 0.02 - 0.015
    Mu = abs(Mu) * 10**5
    Ru = Mu / (0.9*100* ((depth*100)**2))
    As_List = {'RB9' : 0.64 ,'DB10' : 0.785,'DB12' : 1.131,'DB16' : 2.01,'DB20' : 3.14}
    As_List = {'DB10' : 0.785,'DB12' : 1.131,'DB16' : 2.01,'DB20' : 3.14}
    #As_List = {'RB9' : 0.64 ,'DB12' : 1.131,'DB16' : 2.01,'DB20' : 3.14}
    #Centroid_List = {'RB9': 0.0045,'DB10' : 0.005,'DB12':0.006,'DB16':0.008,'DB20':0.01}
    ST_Use = 0
    TypeST = 0
    for i in As_List :
        ST_Use = (As_List[i] * 100 ) / As
        if ST_Use > 7.5 :
            ST_Use = (ST_Use / 100 // 0.025 ) *0.025
            if ST_Use > thickness * 3 :
                ST_Use = thickness * 3
            TypeST = i
            break 
    As_temp = As
    if TypeST == 'RB9' :
        Asmin = 0.0025 * 100 * thickness *100
        p = (0.85*fc/2400) * (1- math.sqrt(1- (2*Ru / (0.85*fc))))
        As = p * 100 * depth * 100
        if As < Asmin :
            As = Asmin
        print(As)
        ST_Use = (0.64 * 100 ) / As
        if ST_Use > 7.5 :
            ST_Use = (ST_Use / 100 // 0.025 ) *0.025
        else :
            As = As_temp
            print('As is ',As) 
            for i in As_List :
                if i != 'RB9' :
                    ST_Use = (As_List[i] * 100 ) / As
                    if ST_Use > 7.5 :
                        ST_Use = (ST_Use / 100 // 0.025 ) *0.025
                        if ST_Use > thickness * 3 :
                            ST_Use = thickness * 3
                        TypeST = i
                        break 

    if ST_Use == 0 :
        return 'เหล็กเสริมไม่พอ'
    else :
        return f'{TypeST} @{ST_Use:.3f}'


def SlabGetAs_or_Add(Mu,fc,fy,botcover,thickness,getAs = False,getAdd = False) :
    As_List = {'RB9' : 0.64 ,'DB10' : 0.785,'DB12' : 1.131,'DB16' : 2.01,'DB20' : 3.14}
    safetyfactor = 0.9
    depth = thickness - botcover
    Mu = abs(Mu) * 10**5 # T/m to kg-cm
    Ru = Mu / (safetyfactor*100* ((depth*100)**2))

    p = (0.85*fc/2400) * (1- math.sqrt(1- (2*Ru / (0.85*fc))))
    Asmin = 0.0025 * 100 * depth*100
    As = p * 100 * depth *100
    if As < Asmin :
        As = Asmin

    ST_Use = 0
    TypeST = None
    for i in As_List :
        ST_Use = (As_List[i] * 100 ) / As
        if ST_Use > 7.5 :
            ST_Use = (ST_Use / 100 // 0.025 ) *0.025
            if ST_Use > thickness * 3 :
                ST_Use = thickness * 3
            TypeST = i
            break

    if TypeST != 'RB9' :
       p = (0.85*fc/fy) * (1- math.sqrt(1- (2*Ru / (0.85*fc))))
       Asmin =   0.0018 * 100 * depth*100
       As = p*100*depth*100
       ST_Use = (As_List[i] * 100 ) / As
       if ST_Use < 7.5 :
            TypeST_notAllowable = TypeST
            for i in As_List :
                if i != 'RB9' and i != TypeST_notAllowable :
                    ST_Use = (As_List[i] * 100 ) / As
                    if ST_Use > 7.5 :
                        ST_Use = (ST_Use / 100 // 0.025 ) *0.025
                        if ST_Use > thickness * 3 :
                            ST_Use = thickness * 3
                        TypeST = i
                        break
    if getAs == True and getAdd == False :                    
        return As
    if getAs == False and getAdd == True :
         if ST_Use == 0 :
            return 'เหล็กในตารางไม่พอ'
         else :
            return f'{TypeST} @{ST_Use:.3f}'





if __name__ == '__main__' :    
    #print(thicknessCheck(Mu = 130,Vu = 1,fc = 210,fy = 4000,botcover = 0.02,thickness=0.1,safetyfactor=0.9))    
    print(Slab_getAs(Mu = 1.06,fc = 210,fy = 4000,botcover = 0.02,thickness=0.10,safetyfactor=0.9))
    #print(get_UseST(8.03,thickness = 0.1,Mu=2.92,fc=240))
    #print(SlabGetAs_or_Add(Mu = 1 ,fc = 210 ,fy = 4000 ,botcover = 0.02 ,thickness = 0.1,getAs=True),SlabGetAs_or_Add(Mu = 1 ,fc = 210 ,fy = 4000 ,botcover = 0.02 ,thickness = 0.1,getAdd=True))
