import win32com.client
import os
import re
import time
import pandas as pd

# cleaned function set up
def cleaned (dataset):
    convert_string = str(dataset)
    string = ' ' + convert_string
    replace_0 = re.sub(' \d{,2}\.', '10)', string)
    replace = re.sub(' \d{,2}\)', '10)', replace_0)
    split_replace = replace.split('10)')
    
    stripped = [i.strip() for i in split_replace]
    
    numbered=[]
    for idx, val in enumerate(stripped):
        if len(stripped) > 1:
            numbered.append(str(idx) + '. ' + val)
        else:
            numbered.append(str(idx+1) + '. ' + val)
        
    if len(numbered) > 1:
        del numbered[0]

    joined = '\n'.join(numbered)
    
    return joined

# Clean corrupted .xls files
file_dir = r"C:\Users\hazman.yusoff\OneDrive - PETRONAS\CodeProjects\Risk Management\Data\LIVE"
file_dir_slash = "C:\\Users\\hazman.yusoff\\OneDrive - PETRONAS\\CodeProjects\\Risk Management\\Data\\LIVE\\"
unhide_dir = [f for f in os.listdir(file_dir) if not f.startswith('~')]

print('Convert files format...\n')

for filename in unhide_dir:
    newfile= os.path.splitext(filename)[0] + ".xlsx"
    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = False
    print('Converting ' + filename + ' --> ' + newfile)
    wb = o.Workbooks.Open(file_dir_slash + filename)
    wb.ActiveSheet.SaveAs(file_dir_slash + newfile, 51)
    os.remove(file_dir_slash + filename)
    o.Application.Quit()
    time.sleep(5)

    
upd_dir = [f for f in os.listdir(file_dir) if not f.startswith('~')]

print('\nRenaming files...\n')

for name in upd_dir:
    df = pd.read_excel (file_dir_slash + name)
    if df.columns[3] == 'Risk ID':
        os.rename(file_dir_slash + name, file_dir_slash + 'Risk Profile.xlsx')
        print(name + ' --> ' + 'Risk Profile.xlsx')
    elif df.columns[3] == 'Existing Mitigation Name':
        os.rename(file_dir_slash + name, file_dir_slash + 'Existing Mitigations.xlsx')
        print(name + ' --> ' + 'Existing Mitigations.xlsx')
    elif df.columns[5] == 'Key Risk Indicator Description':
        os.rename(file_dir_slash + name, file_dir_slash + 'KRI.xlsx')
        print(name + ' --> ' + 'KRI.xlsx')
    elif df.columns[3] == 'New Mitigation Name':
        os.rename(file_dir_slash + name, file_dir_slash + 'New Mitigations.xlsx')
        print(name + ' --> ' + 'New Mitigations.xlsx')


# Data read, clean, save
print('\nCleaning data...\n')
data = pd.read_excel (r'C:\Users\hazman.yusoff\OneDrive - PETRONAS\CodeProjects\Risk Management\Data\LIVE\Risk Profile.xlsx')

data['Causes_cleaned'] = data['Causes'].apply(cleaned)
data['Consequences_cleaned'] = data['Consequences'].apply(cleaned)

data.to_excel(r'C:\Users\hazman.yusoff\OneDrive - PETRONAS\CodeProjects\Risk Management\Data\LIVE\Risk Profile.xlsx', sheet_name = 'Report', index = False)
print('Done\n')

os.system('pause')