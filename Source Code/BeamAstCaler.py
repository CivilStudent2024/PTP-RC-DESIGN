import math
import os
def addStirrup(Vu,fc,fy,width,depth,botcover,safetyfactor,stirType):
    Av_List = {'RB6' : 0.5665486 , 'RB9' : 1.272346 ,'RB12' : 2.26194,'DB10' : 1.5708}
    Av = Av_List[stirType]
    Vu = Vu 
    safetyfactor = safetyfactor
    fc = fc
    fy = fy
    b = width
    d = depth-botcover
    Vn = Vu / safetyfactor
    Vc = (0.53*math.sqrt(fc)*b*d) /1000

    Vs = Vn-Vc
    if Vs < 0 :
        Vs = 0

    #ตรวจสอบระยะห่างมากสุด มี smax , smax2 ใช้ค่าน้อยสุด
    smax_list = []
    smax = (Av*fy) / (0.2*math.sqrt(fc)*b)

    if smax > (Av*fy) / (3.5*b) :
        smax = (Av*fy) / (3.5*b)
        smax_list.append(smax)

    if Vs <= 1.1 * math.sqrt(fc) * b * d / 1000 :
        smax2 = d / 2 
        if smax2 > 60 :
            smax2 == 60
        smax_list.append(smax2)
    elif 1.1 * math.sqrt(fc) * b * d / 1000 <= Vs <= 2.1*math.sqrt(fc)*b*d / 1000 :
        smax2 = d / 4 
        if smax2 > 30 :
            smax2 == 30
        smax_list.append(smax2)


    if Vs <= 2.1*math.sqrt(fc)*b*d /1000 :
        if Vs != 0 :
            s = (Av*fy*d)/1000 / Vs
            smax_list.append(s)

    else :
        smax_list.append('Shear')

    if 'Shear' in smax_list :
        return 'Shear'
    else : 
        return (min(smax_list) / 100 // 0.025) *0.025
    
class BeamAstcal :
    def __init__(self,Mu,fc,fy,width,depth,topcover,botcover,safetyfactor):

        self.beta = 0.85
        self.Mu = abs(Mu) * 10**5 # T/m to kg-cm
        self.fc = fc 
        self.fy = fy #ksc
        self.width = width
        self.depth = depth
        self.topcover = topcover
        self.botcover = botcover
        self.safetyfactor = safetyfactor

        if Mu >= 0 :
            self.d = self.depth - self.botcover 
            self.dd = self.topcover
            self.forcetype = 'Bot-Tension'
        
        if Mu < 0 :
            self.d = self.depth - self.topcover
            self.dd = self.botcover
            self.forcetype = 'Top-Tension'


        self.Mn = self.Mu/self.safetyfactor
        self.Ru = self.Mu / (self.safetyfactor*self.width* (self.d**2))
        self.p = ((0.85*self.fc)/self.fy) * (1- math.sqrt(1- (2*self.Ru / (0.85*self.fc))))
        self.pmin = 14/self.fy
        self.pb = (0.85*self.fc/self.fy) * self.beta * (6120/(6120+fy))
        self.pmax = 0.75*self.pb
        print('p',self.p,'pmax',self.pmax)
        self.type = 'singly'

        self.BotAs = 0
        self.TopAs = 0

    def Singly_Doubly_check(self) :
        #print(self.pmin ,self.p ,self.pmax)
        if self.pmin <= self.p <= self.pmax :
            self.type = 'singly'
            #print('singly beam')
        elif self.p > self.pmax :
            self.type = 'doubly'
            #print('doubly beam')
        elif self.p < self.pmin :
            self.type = 'ขนาดหน้าตัดใหญ่เกินไป'
            #print('ขนาดหน้าตัดใหญ่เกินไป')

    def Ast_cal(self) :
        if self.type == 'singly' or self.type == 'ขนาดหน้าตัดใหญ่เกินไป':
            if self.type == 'singly' :
                Ast = self.p*self.width*self.d #เหล็กเสริมที่ต้องการ หน่วย cm

                Mmax = Ast*self.fy*(self.d-(Ast*self.fy / (2*self.beta*self.fc*self.width))) #ไว้ตรวจสอบ
            elif self.type == 'ขนาดหน้าตัดใหญ่เกินไป' :
                p = self.pmin
                Ast = p*self.width*self.d

            self.BotAs = Ast

        
        elif self.type == 'doubly' :
            As1 = self.pmax*self.width*self.d
            Mn1 = (self.pmax*self.width*self.d*self.fy ) * (self.d - (self.pmax*self.fy*self.width*self.d)/(2*self.beta * self.fc * self.width))
            Mn2 = self.Mn - Mn1
            As2 = Mn2 / (self.fy*(self.d-self.dd))
            Ast = As1+As2

            #ตรวจสอบการครากของเหล็กรับแรงอัด
            a = (As1*self.fy) / (0.85*self.fc*self.width)
            c = a / self.beta
            fs = 6120 * (1-(self.dd/c)) # <= 4000 เหล็กเสริมไม่คราก
            if fs > self.fy : #เหล็กเสริมคราก
                fs = self.fy
            As2 = (As2*self.fy)/(fs)

            self.BotAs = Ast ; self.TopAs = As2

    def analyse(self) :
        self.Singly_Doubly_check()
        self.Ast_cal()

        if self.forcetype == 'Top-Tension' :
            tempBotAs = self.BotAs
            tempTopAs = self.TopAs
            self.TopAs = tempBotAs
            self.BotAs = tempTopAs


    def get_As (self,a) :
        if a == 'bottom' :
            return round(self.BotAs,3)
        elif a == 'top' :
            return round(self.TopAs,3)
        
    def get_Type (self) :
        return self.type

if __name__ == '__main__' :        
    try :
        #พารามิเตอร์
        Mu = 20
        fc =240
        fy=4000
        depth = 50
        width = 30
        topcover = 6
        botcover = 6
        safetyfactor = 0.9
        x = BeamAstcal(Mu = Mu ,fc = fc ,fy = fy , width = width ,depth = depth,topcover = topcover,botcover = botcover ,safetyfactor=0.9)
        x.analyse()
        print('ทดสอบผลการออกแบบคานเหล็กเสริมรับแรงดึง+แรงอัด')
        print('-----------------------------------------')
        print('\tMu (T.M) :',Mu)
        print('\tfc (ksc):',fc)
        print('\tfy (ksc):',fy)
        print('-----------------------------------------')
        print('\tประเภทคาน',x.get_Type())
        print('\tเหล็กเสริมบน',x.get_As(a='top'))
        print('\tเหล็กเสริมล่าง',x.get_As(a='bottom'))
        print('-----------------------------------------')

    except :
        print('Moment')
    
    Vu = 35.55*0.85
    fc = 250
    fy = 2400
    width = 30
    depth = 70
    botcover = 6
    safetyfactor = 0.85
    stirType = 'RB9'
    Stirrup_test = addStirrup(Vu = Vu ,fc = fc ,fy = fy ,width = width ,depth = depth ,botcover = botcover
                              , safetyfactor = safetyfactor ,stirType = stirType)
    try :
        print('ทดสอบออกแบบเหล็กปลอก')
        print('-----------------------------------------')
        print('\tVu (Ton) :',Vu)
        print('\tfc (ksc):',fc)
        print('\tfy (ksc):',fy)
        print('-----------------------------------------')
        print('\tชนิดเหล็กปลอก',stirType)
        print('\tระยะห่าง (m)',f'{Stirrup_test:.3f}')
        print('-----------------------------------------')
    except :
        print('Shear')
    