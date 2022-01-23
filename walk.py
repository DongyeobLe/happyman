#!walk.py
import os, openpyxl, time, bom_by_openpyxl as bbo
import sys
sPath = 'C:\\'; md = sPath[0:1]
wb = openpyxl.Workbook()
sh = wb['Sheet']; i = 3; 
vL = []; tm = time.localtime(time.time());
tdy = time.strftime('%Y%m%d', tm)
for fileObj in os.listdir(sPath):
    if os.path.isdir(sPath + fileObj) or os.path.isfile(sPath + fileObj):
        vL.append(sPath  + fileObj)

for list_in_Root in vL:
    if os.path.isdir(list_in_Root):        
        for foldername, subfolders, filenames in os.walk(list_in_Root):
            print('The current folder is %s' %(foldername))
            for filename in filenames:
                try:
                    file_with_path = os.path.join(foldername, filename)
                    sh.cell(i, 1).value = foldername
                    sh.cell(i, 2).value = filename
                    sh.cell(i, 3).value = os.path.getsize(file_with_path)
                    sh.cell(i, 4).value = time.ctime(os.path.getmtime(file_with_path))
                    sh.cell(i, 5).value = time.ctime(os.path.getctime(file_with_path))
                    sh.cell(i, 6).value = time.ctime(os.path.getatime(file_with_path))
                    i +=  1  
                except  OSError:
                    print('OSError was occured in %s .' %(filename))
    elif os.path.isfile(list_in_Root):
        sh.cell(i, 1).value = sPath
        sh.cell(i, 2).value = os.path.basename(list_in_Root)
        sh.cell(i, 3).value = os.path.getsize(list_in_Root)
        sh.cell(i, 4).value = time.ctime(os.path.getmtime(list_in_Root))
        sh.cell(i, 5).value = time.ctime(os.path.getctime(list_in_Root))
        sh.cell(i, 6).value = time.ctime(os.path.getatime(list_in_Root))
bbo.write_Title([sh, [2,1], [2,6]], ['Path', 'Filename', 'Size(Byte)',\
    'Modified Date', 'Created Date', 'Accessed Date'])  
#bbo.column_width_apply(sh)
if md == 'C':
    wb.save(md +':\\Users\\ldykm\\filelist_in_' + md + '_Root_' + str(tdy) + '.xlsx')
else:
    wb.save(md +':\\filelist_in_' + md + '_Root_' + str(tdy) + '.xlsx')