from tkinter import *
from tkinter.filedialog import askopenfilenames
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
import tkinter as tk
from tkinter import ttk
import pandas as pd
class GUI:
    def __init__(self,master):
        self.master=master
        self.filePaths=None
        self.fileSavePath=None
        self.AnalysisfilePath=None
        self.AnalysisfilePathtwo=None
        self.headerOne=0
        self.headerTwo=0
        self.FilePathButton=Button(text="选择文件",height=1,command=self.filePath)
        self.FilePathButton.place(x=20,y=10)
        self.FileSavePathButton=Button(text="储存目录",height=1,command=self.selectSavePath)
        self.FileSavePathButton.place(x=80,y=10)
        self.FileAddButton=Button(text="合并文件",height=1,command=self.fileAdd)
        self.FileAddButton.place(x=140,y=10)
        self.PathText=Text(bg='lightblue',fg='chocolate',relief='sunken',bd=3,height=20,width=28,padx=1,pady=1,state='normal',cursor='arrow',font=('黑体', 9),wrap='char') 
        self.PathText.place(x=20,y=45)
        
        
        self.FileAnalisisButton=Button(text="选择文件一",height=1,command=self.fileAnalysisPath)
        self.FileAnalisisButton.place(x=250,y=10)
        self.FileAnalisisLabel=Label(text="选择Sheet表:")
        self.FileAnalisisLabel.place(x=325,y=15)
        self.FileAnalisisButton=Button(text="选择文件二",height=1,command=self.fileAnalysisPathTow)
        self.FileAnalisisButton.place(x=250,y=45)
        self.FileAnalisisLabel=Label(text="选择Sheet表:")
        self.FileAnalisisLabel.place(x=325,y=50)
        self.FileReadInfoButton=Button(text="获取文件信息",height=1,command=self.fileReadInfo)
        self.FileReadInfoButton.place(x=250,y=263)
        self.FileOutPathButton=Button(text="选择储存目录",height=1,command=self.selectSavePath)
        self.FileOutPathButton.place(x=335,y=263)
        self.FileAnalisisButton=Button(text="匹配开始",height=1,command=self.AdvFileSelect)                               #.fileAnalysis)
        self.FileAnalisisButton.place(x=420,y=263)
        self.FileAnalisisButton=Button(text="合并文件",height=1,command=self.fileAnalysisAdd)
        self.FileAnalisisButton.place(x=580,y=263)
        self.FileAnalisisLabel=Label(text="选择文件一标表头:")
        self.FileAnalisisLabel.place(x=250,y=80)
        self.FileAnalisisLabel=Label(text="选择文件二标表头:")
        self.FileAnalisisLabel.place(x=250,y=115)
        headNU=[0,1,2,3,4,5,6,7,8,9,10]        
        self.headerOne = ttk.Combobox(master=root, height=5, width=10, state='readonly', cursor='arrow',font=('', 10), values=headNU)
        self.headerOne.place(x=360,y=82)        
        self.headerTwo = ttk.Combobox(master=root,  height=5, width=10, state='readonly', cursor='arrow',font=('', 10), values=headNU)
        self.headerTwo.place(x=360,y=118)
        
        
        tablelabel='''
        表格匹配：
        选择文件一=》选择文件二=》选择sheet表=》选择文件一表头=》
        选择文件二表头=》获取文件信息=》选择文件一匹配项=》
        选择文件二匹配项=》选择储存目录=》匹配开始
        表头代表excel表格的字段行，比如’零件号’
        从表格的上面往下数第一行为‘0’，比如表头在第五行，需要选择‘4’
        '''
        infolabel='''
        有相同表格头的文件用红色方框上的文件合并操作顺序：         
        选择文件=》储存目录=》合并文件
        选择文件处把需要合并的文件全部拉选
        然后选择一个储存目录用来储存合并后的文件
        合并后的文件名：output.xlsx
        默认合并第一个sheet表。
        '''
        PowerByInfo='PowerBy@LOC4'
        self.PowerByInfo=Label(text=PowerByInfo,anchor=W,justify='left')
        self.PowerByInfo.place(x=590,y=470)
        self.programInfo=Label(text=infolabel,anchor=W,justify='left',bg="lightblue")
        self.programInfo.place(x=1,y=300)
        self.programInfotwo=Label(text=tablelabel,anchor=W,justify='left',bg="lightblue")
        self.programInfotwo.place(x=300,y=300)
        
#==================================合并文件=======================================================================       
    def filePath(self):
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
             
#==============================================================================================================            
    def fileAnalysisPath(self):
        
        self.AnalysisfilePath=askopenfilename(title="打开文件1",filetypes=[('excel表格','*.*')])
        if self.AnalysisfilePath!='' and self.AnalysisfilePath!=None:
            df = pd.read_excel(self.AnalysisfilePath,sheet_name=None)
            self.sheetnameOne=[]
            for i in df.keys():           
               self.sheetnameOne.append(i)      
            self.One_sheet_nm = ttk.Combobox(master=root, height=5, width=10, state='readonly', cursor='arrow',font=('', 10), values=self.sheetnameOne,)
            self.One_sheet_nm.place(x=400,y=18)
            
            nameOne=str(self.AnalysisfilePath).split(sep="/")[-1]
            self.FilenameOneLabel=Label(text="文件名:"+nameOne)
            self.FilenameOneLabel.place(x=500,y=18)
        else:
            tk.messagebox.showerror('提示','文件未选择')
    def fileAnalysisPathTow(self):
        
        self.AnalysisfilePathtwo=askopenfilename(title="打开文件2",filetypes=[('excel表格','*.*')])
        if self.AnalysisfilePathtwo!='' and self.AnalysisfilePathtwo!=None:
            df = pd.read_excel(self.AnalysisfilePathtwo,sheet_name=None)
            self.sheetnameTwo=[]
            for i in df.keys():
               
               self.sheetnameTwo.append(i)               
            self.Two_sheet_nm = ttk.Combobox(master=root, height=5, width=10, state='readonly', cursor='arrow', font=('', 10), values=self.sheetnameTwo,)
            self.Two_sheet_nm.place(x=400,y=52)
            
            nameTwo=str(self.AnalysisfilePathtwo).split(sep="/")[-1]
            self.FilenameTwoLabel=Label(text="文件名:"+nameTwo)
            self.FilenameTwoLabel.place(x=500,y=50)
        else:
            tk.messagebox.showerror('提示','文件未选择')
        
    def fileAnalysisAdd(self):
        if self.AnalysisfilePath!='' and self.AnalysisfilePath!=None:
            if self.AnalysisfilePathtwo!='' and self.AnalysisfilePathtwo!=None:
                if self.One_sheet_nm.get()!='' and self.Two_sheet_nm.get()!='': 
                    dataAnalysisOne=pd.read_excel(self.AnalysisfilePath,sheet_name=self.One_sheet_nm.get())
                    dataAnalysisTwo=pd.read_excel(self.AnalysisfilePathtwo,sheet_name=self.Two_sheet_nm.get() )
                    df1 = pd.DataFrame(dataAnalysisOne)
                    df2 = pd.DataFrame(dataAnalysisTwo)
                    result = pd.merge(df1,df2)
                    writer = pd.ExcelWriter(self.fileSavePath+'/result.xlsx',index=False)
                    result.to_excel(writer,index=False)
                    writer.save()
                    tk.messagebox.showinfo('完成','合并完成')
                else:
                    tk.messagebox.showerror('提示','请选择Sheet表单')
            else:
                tk.messagebox.showerror('提示','文件二未选择')
        else:
            tk.messagebox.showerror('提示','文件一未选择')
    def AdvFileSelect(self):
       
        if self.FileColumnsOne_nm.get()!=''and self.FileColumnsTwo_nm.get()!='':
                       
            if self.FileOnecolumns.get()!='' and self.FileTwocolumns.get()!='':
                self.AdvFileAnalysisFour()
            elif self.FileOnecolumns.get()!='' and self.FileTwocolumns.get()=='':
                self.AdvFileAnalysisThree()
            elif self.FileOnecolumns.get()=='' and self.FileTwocolumns.get()!='':
                self.AdvFileAnalysisThreeTwo()
            else:
                self.AdvFileAnalysisTwo()            
        else:
            tk.messagebox.showerror('错误','匹配项一和二必须选择')
    def AdvFileAnalysisTwo(self):
        if self.AnalysisfilePath!=None and self.AnalysisfilePathtwo!=None :
            if self.fileSavePath!=None:
                df1_list=self.df1[self.FileColumnsOne_nm.get()].values.tolist() 
                df2_list=self.df2[self.FileColumnsTwo_nm.get()].values.tolist()
                self.result_list=[]
                for i in df1_list:
                    for n in df2_list :
                        if i==n:
                           
                           self.result_list.append(i)
                saveData=pd.DataFrame(self.result_list)
                saveData.to_excel(excel_writer=self.fileSavePath+'/outer_function.xlsx',index=False)
                tk.messagebox.showinfo('完成','匹配完成')
            else:
                tk.messagebox.showerror('错误','请选择储存目录')
        else:
            tk.messagebox.showerror('错误','文件未指定')

    def AdvFileAnalysisThree(self):
        if self.AnalysisfilePath!=None and self.AnalysisfilePathtwo!=None :
            if self.fileSavePath!=None:
                df1_list=self.df1[self.FileColumnsOne_nm.get()].values.tolist() 
                df2_list=self.df2[self.FileColumnsTwo_nm.get()].values.tolist()
                self.result_list=[]
                self.other_list=[]
                df4=pd.DataFrame()
                for i in df1_list:
                    for n in df2_list :
                        if i==n:
                            
                            df3=self.df1.loc[self.df1[self.FileColumnsOne_nm.get()]==n,[self.FileOnecolumns.get(),self.FileColumnsOne_nm.get()]]                           
                            df4=pd.concat([df4,df3])                            
                            self.result_list.append(i)                             
                saveData=pd.DataFrame(self.result_list)
                df4.to_excel(excel_writer=self.fileSavePath+'/outer_function_two.xlsx',index=False)
                saveData.to_excel(excel_writer=self.fileSavePath+'/outer_function.xlsx',index=False)
                tk.messagebox.showinfo('完成','匹配完成')
            else:
                tk.messagebox.showerror('错误','请选择储存目录')
        else:
            tk.messagebox.showerror('错误','文件未指定')
    def AdvFileAnalysisThreeTwo(self):
        if self.AnalysisfilePath!=None and self.AnalysisfilePathtwo!=None :
            if self.fileSavePath!=None:
                df1_list=self.df1[self.FileColumnsOne_nm.get()].values.tolist() 
                df2_list=self.df2[self.FileColumnsTwo_nm.get()].values.tolist()
                self.result_list=[]
                self.other_list=[]
                df4=pd.DataFrame()
                for one in df1_list:
                    for two in df2_list :
                        if one==two:
                            
                            df3=self.df2.loc[self.df2[self.FileColumnsTwo_nm.get()]==two,[self.FileTwocolumns.get(),self.FileColumnsTwo_nm.get()]]
                                                         
                            df4=pd.concat([df4,df3])                            
                            self.result_list.append(one)                             
                saveData=pd.DataFrame(self.result_list)
                df4.to_excel(excel_writer=self.fileSavePath+'/outer_function_two.xlsx',index=False)
                saveData.to_excel(excel_writer=self.fileSavePath+'/outer_function.xlsx',index=False)
                tk.messagebox.showinfo('完成','匹配完成')
            else:
                tk.messagebox.showerror('错误','请选择储存目录')
        else:
            tk.messagebox.showerror('错误','文件未指定')
    def AdvFileAnalysisFour(self):
        if self.AnalysisfilePath!=None and self.AnalysisfilePathtwo!=None :
            if self.fileSavePath!=None:
                df1_list=self.df1[self.FileColumnsOne_nm.get()].values.tolist() 
                df2_list=self.df2[self.FileColumnsTwo_nm.get()].values.tolist()
                self.result_list=[]
                self.other_list=[]
                df5=pd.DataFrame()
                df6=pd.DataFrame()
                for one in df1_list:
                    for two in df2_list :
                        if one==two:
                            df4=self.df1.loc[self.df1[self.FileColumnsOne_nm.get()]==two,[self.FileOnecolumns.get(),self.FileColumnsOne_nm.get()]] 
                            df3=self.df2.loc[self.df2[self.FileColumnsTwo_nm.get()]==two,[self.FileTwocolumns.get(),self.FileColumnsTwo_nm.get()]]                           
                            df5=pd.concat([df5,df3])
                            df6=pd.concat([df6,df4])                           
                            self.result_list.append(one)
                print(df5)
                print(df6)                
                saveData=pd.DataFrame(self.result_list)
                #df7=pd.merge(df5,df6)
                
                with pd.ExcelWriter(self.fileSavePath+'/outer_function_two.xlsx') as writer:
                    df5.to_excel(writer,index=False)
                    df6.to_excel(writer,index=False,startcol=df5.shape[1])
                saveData.to_excel(excel_writer=self.fileSavePath+'/outer_function.xlsx')
                tk.messagebox.showinfo('完成','匹配完成')
            else:
                tk.messagebox.showerror('错误','请选择储存目录')
        else:
            tk.messagebox.showerror('错误','文件未指定')
    def fileReadInfo(self):
        if self.AnalysisfilePath!='' and self.AnalysisfilePathtwo!='' :
            if self.One_sheet_nm.get()!='' and self.Two_sheet_nm.get()!='':
                headeone=int(self.headerOne.get())
                headetwo=int(self.headerTwo.get())
                if headeone==0 and headetwo==0:
                    tk.messagebox.showerror('提示','默认表头为第一行数据')
                dataAnalysisOne=pd.read_excel(self.AnalysisfilePath,dtype=str,sheet_name=self.One_sheet_nm.get(),header=headeone)
                dataAnalysisTwo=pd.read_excel(self.AnalysisfilePathtwo,dtype=str,sheet_name=self.Two_sheet_nm.get(),header=headetwo)
                self.df1 = pd.DataFrame(dataAnalysisOne)
                self.df2 = pd.DataFrame(dataAnalysisTwo)
                self.FileInfoOneLabel=Label(text="文件一信息：")
                self.FileInfoOneLabel.place(x=460,y=80)
                self.FileInfoTwoLabel=Label(text="文件二信息：")
                self.FileInfoTwoLabel.place(x=460,y=118)
                
                self.FileColOneLabel=Label(text='总行数:'+str(self.df1.shape[0]))
                self.FileColOneLabel.place(x=530,y=80)
                self.FileColTwoLabel=Label(text='总行数:'+str(self.df2.shape[0]))
                self.FileColTwoLabel.place(x=530,y=118)
                
                self.FileRowOneLabel=Label(text='总列数:'+str(self.df1.shape[1]))
                self.FileRowOneLabel.place(x=615,y=80)
                
                self.FileRowTwoLabel=Label(text='总列数:'+str(self.df2.shape[1]))
                self.FileRowTwoLabel.place(x=615,y=118)
                
                self.FileColumnsOneLabel=Label(text="选择一文件匹配项:")
                self.FileColumnsOneLabel.place(x=250,y=143)
                self.FileColumnsTwoLabel=Label(text="选择文件二匹配项:")
                self.FileColumnsTwoLabel.place(x=250,y=175)
                self.FileColumnsOne_nm = ttk.Combobox(master=root, height=5, width=10, state='readonly', cursor='arrow', font=('', 10), values=self.df1.columns.values.tolist())
                self.FileColumnsOne_nm.place(x=360,y=146)
                self.FileColumnsTwo_nm = ttk.Combobox( master=root, height=5, width=10, state='readonly', cursor='arrow', font=('', 10), values=self.df2.columns.values.tolist())
                self.FileColumnsTwo_nm.place(x=360,y=178)
                
                self.FileOneTwoLabel=Label(text="匹配项二:")
                self.FileOneTwoLabel.place(x=460,y=143)
                self.FileOnecolumns=ttk.Combobox(master=root,  height=5, width=10, state='readonly', cursor='arrow',font=('', 10), values=self.df1.columns.values.tolist())
                self.FileOnecolumns.place(x=550,y=146)
                
                self.FileTwoTwoLabel=Label(text="匹配项二:")
                self.FileTwoTwoLabel.place(x=460,y=175)
                self.FileTwocolumns=ttk.Combobox(master=root,  height=5, width=10, state='readonly', cursor='arrow',font=('', 10), values=self.df2.columns.values.tolist())
                self.FileTwocolumns.place(x=550,y=178)
            else:
                tk.messagebox.showerror('提示','请选择Sheet表单')
        else:
            tk.messagebox.showerror('提示','文件未选择')
if __name__ == "__main__":
    root = Tk()
    root.title("excel文件小工具")

    frame = Frame(root)

    root.geometry("700x500")
    app = GUI(frame)

    root.mainloop()