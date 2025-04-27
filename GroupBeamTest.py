import pandas as pd
filepath = r'C:\ProjectVenv\.venv\ProjectHouse\Exhouse-Analysed.xlsx'
GB_list = {}
GB_list_result = {}
beam_add_check = []
count_group = 0
for i in range(1,11) :
    if i not in beam_add_check :
        count_group += 1
        GB_list[f"B{count_group}"] = [f'B{i}']
        beam1 = pd.read_excel(filepath,sheet_name=f'Beam{i}') 
        beam1 = beam1[['Span No','Top-As','Bot-As','RB6','RB9']]
        for j in range(1,11) :
            if j != i :
                beam2 = pd.read_excel(filepath,sheet_name=f'Beam{j}')
                beam2 = beam2[['Span No','Top-As','Bot-As','RB6','RB9']]
                if  beam1.shape == beam2.shape and (beam1 == beam2).all().all() == True :
                    print('Group',i,j)
                    GB_list[f"B{count_group}"].append(f"B{j}")
                    beam_add_check.append(j)
        GB_list_result[f"B{count_group}"] = ','.join(GB_list[f"B{count_group}"])

print(GB_list_result)
