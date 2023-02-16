import tkinter
from tkinter import *
from tkinter.filedialog import askopenfilenames
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
import tkinter as tk
import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
class GUI:
    def __init__(self,master):
        self.master=master
        self.filePaths=None
        self.fileSavePath=None
        self.AnalysisfilePath=None
        self.AnalysisfilePathtwo=None
        self.headerOne=0
        self.headerTwo=0
        self.FilePathButton=ttk.Button(text="选择文件",command=self.filePath)
        self.FilePathButton.place(x=20,y=10)
        self.FileSavePathButton=ttk.Button(text="储存目录",command=self.selectSavePath)
        self.FileSavePathButton.place(x=95,y=10)
        self.FileAddButton=ttk.Button(text="合并为单sheet",command=self.fileAdd)
        self.FileAddButton.place(x=20,y=45)
        self.FileAddButton=ttk.Button(text="合并为多sheet",command=self.sheetAdd)
        self.FileAddButton.place(x=126,y=45)
        self.FileAddButton=ttk.Button(text="多表格交集匹配",command=self.sheetMerge)
        self.FileAddButton.place(x=232,y=45)                
        self.PathText=ttk.Text(bg='lightblue',fg='chocolate',relief='sunken',bd=3,height=16,width=57,padx=1,pady=1,state='normal',cursor='arrow',font=('黑体', 9),wrap='char') 
        self.PathText.place(x=20,y=80)
    def filePath(self):
        self.PathText.delete(0.0, 'end')
        self.filePaths=askopenfilenames(title="打开文件",filetypes=[('excel表格','*.*')])
        if self.filePaths!='':
            count=len(self.filePaths)
            
            for i in range(count):
                name=str(self.filePaths[i]).split(sep="/")[-1]
                
                self.PathText.insert('insert','\n')
                self.PathText.insert('insert',name)
        else:
            tk.messagebox.showerror('提示','文件未选择')
    def selectSavePath(self):
        self.fileSavePath=askdirectory(title="选择文件储存目录")
        if self.fileSavePath=='':
            tk.messagebox.showerror('提示','储存目录未选择')
    def fileAdd(self):
        if self.filePaths!=None and self.fileSavePath!=None :
            frames = []
            for i in range(len(self.filePaths)):
                
                frames.append(pd.read_excel(self.filePaths[i],dtype=str))
            writer = pd.ExcelWriter(self.fileSavePath+'/output.xlsx')
            pd.concat(frames).to_excel(writer,'Sheet1',index=False)
            writer.save()
            tk.messagebox.showinfo('完成','合并完成')
        else :
             tk.messagebox.showerror('错误','请选择文文件或文件储存目录')
    def sheetAdd(self):
        if self.filePaths!=None and self.fileSavePath!=None :
            frames = []
            x=0
            writer = pd.ExcelWriter(self.fileSavePath+'/output.xlsx')
            for i in range(len(self.filePaths)):
                x=x+1
                y=str(x)
                z=pd.read_excel(self.filePaths[i],dtype=str)
                
                z.to_excel(writer,'Sheet'+y,index=False)
                
            tk.messagebox.showinfo('完成','合并完成')
            writer.save()
        else :
             tk.messagebox.showerror('错误','请选择文文件或文件储存目录')
    def sheetMerge(self):
        self.tableCol = tkinter.simpledialog.askstring(title = '信息',prompt='请输入匹配字段名：',initialvalue = None)
        if self.filePaths!=None and self.fileSavePath!=None :
            writer =self.fileSavePath+'/output.xlsx'
            self.outData=pd.DataFrame()
            for i in range(len(self.filePaths)):
                z=pd.read_excel(self.filePaths[i],dtype=str)
                if self.outData.empty:
                    self.outData=z
                else:
                    self.outData=pd.merge(self.outData,z,how='inner',on=self.tableCol)
            self.outData.to_excel(writer,index=False)             
            tk.messagebox.showinfo('完成','合并完成')
        else :
             tk.messagebox.showerror('错误','请选择文文件或文件储存目录')              
if __name__ == "__main__":
    root = Tk()
    root.title("excel文件合并")

    frame = Frame(root)

    root.geometry("400x320")
    app = GUI(frame)

    root.mainloop()  
