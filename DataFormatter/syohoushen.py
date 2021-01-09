import tkinter
import traceback
import sys
import tkinter.messagebox as msg
import y_ALL
import ippanmeishohoumaster
import relation
import tp_01_04
import tp_05
import os

if __name__ == "__main__":
    try:
        # Check if directory exists, if not, create it
        dir1=os.path.isdir('resource')
        dir2=os.path.isdir('templates')
        
        if not dir1 or not dir2:
            os.makedirs('resource')
            os.makedirs('templates')

        # tkinter INIT
        root = tkinter.Tk()
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        root.title("INSERT_SQLを作成するファイルを選んでください。")
        root.geometry("500x130+"+str(int(screen_width/2)-200)+"+"+str(int(screen_height/2)-200))
        tkinter.Button(root,text="y_ALL*.csv",width=190,command=y_ALL.start).pack()
        tkinter.Button(root,text="ippanmeishohoumaster_*.xlsx",width=190,command=ippanmeishohoumaster.start).pack()
        tkinter.Button(root,text="relation",width=190,command=relation.start).pack()
        tkinter.Button(root,text="tp*-01_05.xlsx",width=190,command=tp_01_04.start).pack()
        tkinter.Button(root,text="tp-01_04.csv",width=190,command=tp_05.start).pack()

        root.mainloop()
    except Exception as ex:
        print(traceback.print_exc())
        msg.showerror(title='エラー', message='エラー履歴を確認してください。')