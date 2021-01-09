import tkinter as tk
import tkinter.filedialog as fl
import tkinter.messagebox as msg
import csv
import math
import traceback
import sys

def readCsv():
    '''READ_CSV'''
    fileName=fl.askopenfilename(initialdir="./resource",filetypes=[('csv file', '.csv')],title="select [y_ALL*.csv ]file")
    row_All=[]
    qeuryArr=[]

    #FILE OPEN
    with open(fileName,'r') as f:
        reader = csv.reader(f) 
        for row in reader:
            add_comma="','".join(row)
            row_All.append("('"+add_comma+"')")

    #CREATE LIST
    while True: 
        num=len(qeuryArr)+1
        rowcnt_devided=math.ceil(len(row_All)/3000)-1

        #qeuryArr.append(row_All[(0*num):(3000*num)])3000*num
        add_3000=row_All[(num-1)*3000:num*3000]
        add_3000_comma=",".join(add_3000)
        qeuryArr.append("INSERT INTO `medicine_mst` VALUES "+add_3000_comma+";\n")

        if rowcnt_devided<num:              
            break
        
    return qeuryArr

# def writeCsv(qeuryArr):
#     '''WRITE_CSV'''
#     with open('./templates/sos_medicine_mst_insert.sql', 'w',-1,'utf-8') as f:
#         for row in qeuryArr:
#             f.write(row)

def combineSql(qeuryArr):
    '''FIRST+MAIN+LAST SQL'''
    first_stream=""
    main_stream=''.join(qeuryArr)
    last_stream=""

    #FIRST FILE
    fileName_first=fl.askopenfilename(initialdir="./templates",filetypes=[('sql file', 'sos_medicine_mst_template_first.sql')],title="select [sos_medicine_mst_template_first.sql] file")
    with open(fileName_first, 'r', encoding="utf-8") as firstFile:
        first_stream = firstFile.read()

    # #MAIN FILE
    # with open("./templates/sos_medicine_mst_insert.sql", 'r',encoding="utf-8") as mainFile:
    #     main_stream = mainFile.read()

    #LAST FILE
    fileName_last=fl.askopenfilename(initialdir="./templates",filetypes=[('sql file', 'sos_medicine_mst_template_last.sql')],title="select [sos_medicine_mst_template_last.sql] file")
    with open(fileName_last, 'r', encoding="utf-8") as lastFile:
        last_stream = lastFile.read()

    with open('./sos_medicine_mst.sql', 'w',-1,'utf-8') as f:
        f.write(first_stream+"\n")
        f.write(main_stream+"\n")
        f.write(last_stream+"\n")

def start():
    try:
        combineSql(readCsv())
        msg.showinfo(title='成功', message='保存しました。')
    except FileNotFoundError:
        pass
    except Exception: 
        print(traceback.print_exc())
        msg.showerror(title='エラー', message='エラー履歴を確認してください。')

