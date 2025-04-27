import os
from PyPDF2 import PdfMerger
folder_path = os.getcwd()+'\Lanna-C3-5507'
Lanna_all = []
merge = PdfMerger()

for i in range(1,12) :
    if i < 10 :
        globals()[f'T20{i}'] = folder_path+f'\LANNA-T20{i}-C3-5507.pdf'
        globals()[f'MT20{i}'] = folder_path+f'\T20{i}.pdf'
        Lanna_all.extend([globals()[f'MT20{i}'],globals()[f'T20{i}']])
    else :
        globals()[f'T21{i-10}'] = folder_path+f'\LANNA-T21{i-10}-C3-5507.pdf'
        globals()[f'MT21{i-10}'] = folder_path+f'\T21{i-10}.pdf'        
        Lanna_all.extend([globals()[f'MT21{i-10}'],globals()[f'T21{i-10}']])


Lanna_all.append(folder_path+f'\T3.pdf')
Lanna_all.append(folder_path+f'\LANNA-T3-C3-5507.pdf')
Lanna_all.append(folder_path+f'\T1.pdf')
Lanna_all.append(folder_path+f'\LANNA-T1-C3-5507.pdf')

print(Lanna_all)
for j in Lanna_all :
    merge.append(j)
merge.write('Lanna-C3-5507\LANNA-C3-5507.pdf')
merge.close()
