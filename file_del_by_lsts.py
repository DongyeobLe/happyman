#!walk.py
def filename_endwith(strA, lsts):
    for i in range(len(lsts)):
        if strA.endswith(lsts[i]):
            return True
    return False

def wr_in_sht(sht, iC, aset):
    for i in range(1, len(aset) + 1):
        sht.cell(i, iC).value = aset.pop()
        i += 1
           
import os, openpyxl, shutil, send2trash
sPath = 'C:\\Py\\UT' ; wPath = 'C:\\Py\\UT\\ExcelFiles'
delextLst = ['bin', 'pdf', 'dwg','tpl', 'nc1','lay','dg','css' ,'dll','html','gif','zsol', \
            'log', 'db1','db', 'dwl2', 'xml','inp','dwl','ini','txt']
wb = openpyxl.Workbook()
sh = wb['Sheet']
os.makedirs(wPath , exist_ok = True)
icnt = 0; i = 1; s = set();
for foldername, subfolders, filenames in os.walk(sPath):
    print('The current folder is %s' %(foldername))
    for filename in filenames:
        originalfile = os.path.join(foldername, filename)
        newfile = os.path.join(wPath, filename)
        if filename_endwith(filename, ['xlsx', 'xlsm','xlsb','xls']):
            try:
                shutil.copy(originalfile, newfile)
            except shutil.SameFileError as e:
                print(f'{filename} is already existing in the {foldername} ')    
        elif filename_endwith(filename, delextLst):    
            send2trash.send2trash(os.path.join(foldername, filename))
            icnt += 1
            print('Total %s files are being deleted.' %(icnt))
        elif  not '.' in filename:
            send2trash.send2trash(os.psth.join(foldername, filename))
            icnt += 1
            print('Total %s files are being deleted.' %(icnt))
        else:
            ext = os.path.splitext(os.path.join(foldername, filename))
            s.add(ext[1])
wr_in_sht(sh, 1, s)
wb.save(sPath + '\\ext_lists.xlsx')
   
for foldername, subfolders, filenames, in os.walk(sPath):
    try:
        os.rmdir(foldername)
        print('%s is being deleted.' %(foldername)) 
    except OSError as e:
        print(f'Exception occured in {foldername} : {e.strerror}')
    print('Empty folder deleting was all done.')                       
