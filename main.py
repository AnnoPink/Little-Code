import os,sys
import xlwt
import xlrd
import shutil

data = xlrd.open_workbook('class/table.xlsx')
table = data.sheets()[0]
col = table.col_values(2)
cell_2_11 = table.cell(2,11).value  #某个单元格的值，从0开始计数

col_L = table.col(11)  #提取出L列的所有数据
col_C = table.col(2)
newcell_2_11= col_L[2].value   #提取出（2，11）的组别
filename1=[]  # 存放"地理学"类文件名称

for i in range(2,114):
    if col_L[i].value == newcell_2_11 :
        filename1.append(col_C[i].value)
print(data)
print(table)
print(col) #col 列
print(cell_2_11)
print(col_L)
print(newcell_2_11)
print(filename1)

#测试部分，一篇文章
testname = filename1[0]
print(testname)

path = 'class/paper'
path_new='class/地理学（含地理、资源、环境）/'
filelist=[]
for root,dirs,files in os.walk(path):
    for name in files:
        if testname in name:
            print(name)
            full_path = os.path.join(root,name)
            despath = path_new + name
            shutil.move(full_path,despath)
            break
        #fitfile=filelist.append(os.path.join(path,name))
        #print(os.path.join(path,name))

for i in filelist:
    if os.path.isfile(i):
        print(i)
        if testname in os.path.split(i)[1]:
            print(i)


#找出整个地理学的
#testname = filename1[0]
print(testname)

path = 'class/paper'
path_new='class/地理学（含地理、资源、环境）/'
filelist=[]
for testname in filename1:
    for root,dirs,files in os.walk(path):
        for name in files:
            if testname in name:
                print(name)
                full_path = os.path.join(root,name)
                despath = path_new + name
                shutil.move(full_path,despath)
                break
        #fitfile=filelist.append(os.path.join(path,name))
        #print(os.path.join(path,name))

for i in filelist:
    if os.path.isfile(i):
        print(i)
        if testname in os.path.split(i)[1]:
            print(i)

#找出整个教育学的
#testname = filename1[0]
newcell_10_11= col_L[10].value   #提取出（2，11）的组别
print(newcell_10_11)
filename1=[]  # 存放"教育学"类文件名称

for i in range(2,114):
    if col_L[i].value == newcell_10_11 :
        filename1.append(col_C[i].value)

#print(testname)

path = 'class/paper'
path_new='class/教育学（含教育、心理、体育）/'
filelist=[]
for testname in filename1:
    for root,dirs,files in os.walk(path):
        for name in files:
            if testname in name:
                print(name)
                full_path = os.path.join(root,name)
                despath = path_new + name
                shutil.move(full_path,despath)
                break
        #fitfile=filelist.append(os.path.join(path,name))
        #print(os.path.join(path,name))

for i in filelist:
    if os.path.isfile(i):
        print(i)
        if testname in os.path.split(i)[1]:
            print(i)

#找出整个文学的
#testname = filename1[0]
newcell_3_11= col_L[3].value   #提取出（2，11）的组别
print(newcell_3_11)
filename1=[]  # 存放"教育学"类文件名称

for i in range(2,114):
    if col_L[i].value == newcell_3_11 :
        filename1.append(col_C[i].value)

#print(testname)

path = 'class/paper'
path_new='class/文学/'
filelist=[]
for testname in filename1:
    for root,dirs,files in os.walk(path):
        for name in files:
            if testname in name:
                print(name)
                full_path = os.path.join(root,name)
                despath = path_new + name
                shutil.move(full_path,despath)
                break
        #fitfile=filelist.append(os.path.join(path,name))
        #print(os.path.join(path,name))

for i in filelist:
    if os.path.isfile(i):
        print(i)
        if testname in os.path.split(i)[1]:
            print(i)

#找出整个哲学的
#testname = filename1[0]
newcell_4_11= col_L[4].value   #提取出（2，11）的组别
print(newcell_4_11)
filename1=[]  # 存放"哲学"类文件名称

for i in range(2,114):
    if col_L[i].value == newcell_4_11 :
        filename1.append(col_C[i].value)

#print(testname)

path = 'class/paper'
path_new='class/哲学/'
filelist=[]
for testname in filename1:
    for root,dirs,files in os.walk(path):
        for name in files:
            if testname in name:
                print(name)
                full_path = os.path.join(root,name)
                despath = path_new + name
                shutil.move(full_path,despath)
                break
        #fitfile=filelist.append(os.path.join(path,name))
        #print(os.path.join(path,name))

for i in filelist:
    if os.path.isfile(i):
        print(i)
        if testname in os.path.split(i)[1]:
            print(i)



#找出整个历史学的
#testname = filename1[0]
newcell_5_11= col_L[5].value   #提取出（2，11）的组别
print(newcell_5_11)
filename1=[]  # 存放"历史学"类文件名称

for i in range(2,114):
    if col_L[i].value == newcell_5_11 :
        filename1.append(col_C[i].value)

#print(testname)

path = 'class/paper'
path_new='class/历史学/'
filelist=[]
for testname in filename1:
    for root,dirs,files in os.walk(path):
        for name in files:
            if testname in name:
                print(name)
                full_path = os.path.join(root,name)
                despath = path_new + name
                shutil.move(full_path,despath)
                break
        #fitfile=filelist.append(os.path.join(path,name))
        #print(os.path.join(path,name))

for i in filelist:
    if os.path.isfile(i):
        print(i)
        if testname in os.path.split(i)[1]:
            print(i)

#社会学（含社会、经济、政管、统计、法学）
#testname = filename1[0]
newcell_15_11= col_L[15].value   #提取出（2，11）的组别
print(newcell_15_11)
filename1=[]  # 存放"社会学（含社会、经济、政管、统计、法学）"类文件名称

for i in range(2,114):
    if col_L[i].value == newcell_15_11 :
        filename1.append(col_C[i].value)

#print(testname)

path = 'class/paper'
path_new='class/社会学（含社会、经济、政管、统计、法学）/'
filelist=[]
for testname in filename1:
    for root,dirs,files in os.walk(path):
        for name in files:
            if testname in name:
                print(name)
                full_path = os.path.join(root,name)
                despath = path_new + name
                shutil.move(full_path,despath)
                break
        #fitfile=filelist.append(os.path.join(path,name))
        #print(os.path.join(path,name))

for i in filelist:
    if os.path.isfile(i):
        print(i)
        if testname in os.path.split(i)[1]:
            print(i)

#通识必修（含思政、军理）
#testname = filename1[0]
newcell_9_11= col_L[9].value   #提取出（2，11）的组别
print(newcell_9_11)
filename1=[]  # 存放"通识必修（含思政、军理）"类文件名称

for i in range(2,114):
    if col_L[i].value == newcell_9_11 :
        filename1.append(col_C[i].value)

#print(testname)

path = 'class/paper'
path_new='class/通识必修（含思政、军理）/'
filelist=[]
for testname in filename1:
    for root,dirs,files in os.walk(path):
        for name in files:
            if testname in name:
                print(name)
                full_path = os.path.join(root,name)
                despath = path_new + name
                shutil.move(full_path,despath)
                break
        #fitfile=filelist.append(os.path.join(path,name))
        #print(os.path.join(path,name))

for i in filelist:
    if os.path.isfile(i):
        print(i)
        if testname in os.path.split(i)[1]:
            print(i)

#艺术
#testname = filename1[0]
newcell_14_11= col_L[14].value   #提取出（2，11）的组别
print(newcell_14_11)
filename1=[]  # 存放"艺术"类文件名称

for i in range(2,114):
    if col_L[i].value == newcell_14_11 :
        filename1.append(col_C[i].value)

#print(testname)

path = 'class/paper'
path_new='class/艺术/'
filelist=[]
for testname in filename1:
    for root,dirs,files in os.walk(path):
        for name in files:
            if testname in name:
                print(name)
                full_path = os.path.join(root,name)
                despath = path_new + name
                shutil.move(full_path,despath)
                break
        #fitfile=filelist.append(os.path.join(path,name))
        #print(os.path.join(path,name))

for i in filelist:
    if os.path.isfile(i):
        print(i)
        if testname in os.path.split(i)[1]:
            print(i)