import tkinter as tk
from tkinter import filedialog,font
#from algorithm import create_table,sap
import pandas as pd
from ex import *

class MainView(tk.Frame):
    def __init__(self, *args, **kwargs):
        tk.Frame.__init__(self, *args, **kwargs)

        lbli= tk.Label(self, text='Input', fg="blue", font= font.Font(size=8))
        lbli.place(x=10, y=20)
        
        lblo= tk.Label(self, text='Output', fg="blue", font= font.Font(size=8))
        lblo.place(x=10, y=190)

        lbl1= tk.Label(self, text='IdM_Auszug *', font= font.Font(size=11))
        lbl1.place(x=10, y=45)
        e1= tk.Entry(self)
        e1.insert (0, "Datei" )
        e1.place(x=200, y=45, width=170)
        b1= tk.Button(self, text ='Auswählen',bg="#C9C0BB", command =lambda: self.select_file(e1,1))
        b1.place(x=380, y=42)
        
        global ef
        #App = tk.StringVar()
        lblf= tk.Label(self, text='Application', font= font.Font(size=11))
        lblf.place(x=10, y=65)
        ef= tk.Entry(self)
        ef.insert (0, "Application name" )
        ef.place(x=200, y=65, width=170)

        lbl2= tk.Label(self, text='Stellenkonzept *', font= font.Font(size=11))
        lbl2.place(x=10, y=95)
        e2= tk.Entry(self)
        e2.insert (0, "Datei" )
        e2.place(x=200, y=95, width=170)
        b2= tk.Button(self, text ='Auswählen',bg="#C9C0BB", command =lambda: self.select_file(e2,2))
        b2.place(x=380, y=92)
        
        lbl3= tk.Label(self, text='Rollenkonzept *', font= font.Font(size=11))
        lbl3.place(x=10, y=145)
        e3= tk.Entry(self)
        e3.insert (0, "Datei" )
        e3.place(x=200, y=145, width=170)
        b3= tk.Button(self, text ='Auswählen',bg="#C9C0BB", command =lambda: self.select_file(e3,3))
        b3.place(x=380, y=142)
    
        lbl4= tk.Label(self, text='Ordner für Ausgabe *', font= font.Font(size=11))
        lbl4.place(x=10, y=215)
        e4= tk.Entry(self)
        e4.insert (0, "Ordner" )
        e4.place(x=200, y=215, width=170)
        b4= tk.Button(self, text ='Auswählen',bg="#C9C0BB", command =lambda: self.result_dest(e4))
        b4.place(x=380, y=212)

        b5= tk.Button(self, text='Start', font= font.Font(size=13),bg="#C9C0BB", command= self.run)
        b5.place(x=500,y=260, width=80, height=30)

    def select_file(self, x, index):
        x.delete(0,'end')
        file_path= filedialog.askopenfilename(title = "Select A File")
        x.insert(0,file_path)
        #check in which button
        if(index==1):
            global IdM_Auszug
            IdM_Auszug = sap.data_frame(file_path)            
        if(index==2):
            global Stellenkonzept
            Stellenkonzept = sap.data_frame(file_path)
        if(index==3):
            global Rollenkonzept
            Rollenkonzept = sap.data_frame(file_path)

    def result_dest(self,x):
        global directory
        x.delete(0,'end')
        directory = filedialog.askdirectory()
        x.insert(0,directory)

    def run(self):
            arr = ['/Users nicht in Stellenkonzept.xlsx','/Users nicht in IdM_Auszug.xlsx',
            '/Rollen nicht in IDM.xlsx','/Rollen nicht in Rollenkonzept.xlsx','/Übersicht Tabelle.xlsx']
            path = [0]*5
            for i in range(5):
                path[i]= directory+arr[i]

            App= ef.get()
            table1, table2 = sap.step_1(IdM_Auszug,Stellenkonzept,App)
            table3, table4 = sap.step_2(IdM_Auszug,Stellenkonzept,Rollenkonzept,App)
            table5 = sap.step_3(IdM_Auszug,Stellenkonzept,App)
            table1.to_excel(path[0])
            table2.to_excel(path[1])
            table3.to_excel(path[2])
            table4.to_excel(path[3])
            table5.to_excel(path[4])



            columns1=len(table1.columns)
            writer= pd.ExcelWriter(path[0], engine= 'xlsxwriter')
            table1.to_excel(writer, startrow=0,startcol=0, sheet_name= 'sheet1')

            workbook= writer.book
            worksheet= writer.sheets['sheet1']

            format1= workbook.add_format({
                'bg_color': '#0072bb',
                'border': 1,
                'font_color': 'white'})
            worksheet.conditional_format(0,1,0,columns1+2, {'type':   'no_blanks',
                                                'format': format1})

            worksheet.autofilter(0,1,0,columns1)
            worksheet.set_column(0, 10, 20)
            writer.save()

            columns2=len(table2.columns)
            writer= pd.ExcelWriter(path[1], engine= 'xlsxwriter')
            table2.to_excel(writer, startrow=0,startcol=0, sheet_name= 'sheet1')

            workbook= writer.book
            worksheet= writer.sheets['sheet1']

            format1= workbook.add_format({
                'bg_color': '#0072bb',
                'border': 1,
                'font_color': 'white'})
            worksheet.conditional_format(0,1,0,columns2+2, {'type':   'no_blanks',
                                                'format': format1})

            worksheet.autofilter(0,1,0,columns2)
            worksheet.set_column(0, 10, 20)
            writer.save()

            columns3=len(table3.columns)
            writer= pd.ExcelWriter(path[2], engine= 'xlsxwriter')
            table3.to_excel(writer, startrow=0,startcol=0, sheet_name= 'sheet1')

            workbook= writer.book
            worksheet= writer.sheets['sheet1']

            format1= workbook.add_format({
                'bg_color': '#0072bb',
                'border': 1,
                'font_color': 'white'})
            worksheet.conditional_format(0,1,0,columns3+2, {'type':   'no_blanks',
                                                'format': format1})

            worksheet.autofilter(0,1,0,columns3)
            worksheet.set_column(0, 10, 20)
            writer.save()

            columns4=len(table4.columns)
            writer= pd.ExcelWriter(path[3], engine= 'xlsxwriter')
            table4.to_excel(writer, startrow=0,startcol=0, sheet_name= 'sheet1')

            workbook= writer.book
            worksheet= writer.sheets['sheet1']

            format1= workbook.add_format({
                'bg_color': '#0072bb',
                'border': 1,
                'font_color': 'white'})
            worksheet.conditional_format(0,1,0,columns4+2, {'type':   'no_blanks',
                                                'format': format1})

            worksheet.autofilter(0,1,0,columns4)
            worksheet.set_column(0, 10, 20)
            writer.save()

            columns5=len(table5.columns)
            writer= pd.ExcelWriter(path[4], engine= 'xlsxwriter')
            table5.to_excel(writer, startrow=0,startcol=0, sheet_name= 'sheet1')

            workbook= writer.book
            worksheet= writer.sheets['sheet1']

            format1= workbook.add_format({
                'bg_color': '#0072bb',
                'border': 1,
                'font_color': 'white'})
            worksheet.conditional_format(0,1,0,columns5+2, {'type':   'no_blanks',
                                                'format': format1})

            worksheet.autofilter(0,1,0,columns5)
            worksheet.set_column(0, 10, 20)
            writer.save()











            label_end = tk.Label(self, text="files saved", fg="green", font= font.Font(size=15))
            label_end.place(x=270,y=260)

if __name__ == "__main__":
    root = tk.Tk()
    main = MainView(root)
    main.pack(side="top", fill="both", expand=True)
    root.wm_geometry("600x300")
    root.title("SAP Stellenkonzept")
    root.mainloop()
