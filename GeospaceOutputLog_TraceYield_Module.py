#Front End
import os
import sys
from tkinter import*
import tkinter.messagebox
import GeoSpaceOutputLog_BackEnd
import tkinter.ttk as ttk
import tkinter as tk
import sqlite3
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilenames
from tkinter import simpledialog
import pandas as pd
import openpyxl
import csv
import time
import datetime
from datetime import datetime
import matplotlib.pyplot as plt

def TraceYieldReport():
    window = Tk()
    window.title ("Geomerge Output Log Trace Yield Analysis")
    window.geometry("1230x712+10+0")
    window.config(bg="cadet blue")
    window.resizable(0, 0)
    DataFrameTOP = LabelFrame(window, bd = 2, width = 1348, height = 25, padx= 0, pady= 1,relief = RIDGE,
                                       bg = "cadet blue",font=('aerial', 12, 'bold'))
    DataFrameTOP.pack(side=TOP)

    ### Table Define
    TableMargin = Frame(window, bd = 2, padx= 5, pady= 5, relief = RIDGE)
    TableMargin.pack(side=TOP)
    scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
    scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
    tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5",
                                             "column6", "column7", "column8", "column9", "column10"), height=27, show='headings')
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)
    tree.heading("#1", text="File Number", anchor=W) 
    tree.heading("#2", text="Shot Overrides", anchor=W)
    tree.heading("#3", text="Process Type", anchor=W)
    tree.heading("#4", text="Line Number", anchor=W)        
    tree.heading("#5", text="Station Number", anchor=W)
    tree.heading("#6", text="Comments", anchor=W)
    tree.heading("#7", text="Num Seis Channels", anchor=W)        
    tree.heading("#8", text="Num Miss Channels" ,anchor=W)
    tree.heading("#9", text="Num Dead Channels", anchor=W)
    tree.heading("#10", text="% Missing Per Source", anchor=W)        
       
    tree.column('#1', stretch=NO, minwidth=0, width=100)            
    tree.column('#2', stretch=NO, minwidth=0, width=120)
    tree.column('#3', stretch=NO, minwidth=0, width=100)
    tree.column('#4', stretch=NO, minwidth=0, width=100)
    tree.column('#5', stretch=NO, minwidth=0, width=110)
    tree.column('#6', stretch=NO, minwidth=0, width=100)
    tree.column('#7', stretch=NO, minwidth=0, width=130)
    tree.column('#8', stretch=NO, minwidth=0, width=140)
    tree.column('#9', stretch=NO, minwidth=0, width=140)
    tree.column('#10', stretch=NO, minwidth=0, width=150)

    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", font=('Verdana', 8), foreground="black")
    style.configure("Treeview", foreground='black')
    style.configure("Treeview.Heading",font=('Calibri', 11,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')
    tree.pack()

    DataFrameBOTTOM_ACTIONS = LabelFrame(window, bd = 2, width = 1230, height = 8, padx= 0, pady= 2,relief = RIDGE,
                                       bg = "cadet blue",font=('aerial', 12, 'bold'))
    DataFrameBOTTOM_ACTIONS.pack(side=BOTTOM)

    ### Functions
    def ExportTraceReport():
        conn = sqlite3.connect("GeoSpaceOutputLogTraceYield.db")
        Complete_df = pd.read_sql_query("SELECT * FROM GeoSpaceOutputLog ORDER BY `FileNumber` ASC ;", conn)
        conn.commit()
        conn.close()
        ListBox_DF = pd.DataFrame(Complete_df)
        TotalListBox= len(ListBox_DF)
        if TotalListBox >0:    
            Valid_No_Duplicated_OutputLog_DF = pd.DataFrame(Complete_df)
            Valid_No_Duplicated_OutputLog_DF = Valid_No_Duplicated_OutputLog_DF.reset_index(drop=True)
            Length_No_Duplicated_OutputLog  = len(Valid_No_Duplicated_OutputLog_DF)
            SUM_Trace_Recorded_OutputLog    = ((Valid_No_Duplicated_OutputLog_DF['NumSeisChannels']).sum())
            SUM_Trace_Dead_OutputLog        = ((Valid_No_Duplicated_OutputLog_DF['NumMissChannels']).sum())
            Percent_Dead_OutputLog          = round((((SUM_Trace_Dead_OutputLog)/(SUM_Trace_Recorded_OutputLog))*100),2)
            Percent_Yield_OutputLog         = round((100 - Percent_Dead_OutputLog),2)   
            
            TotalUniqueFFID    = Length_No_Duplicated_OutputLog
            TotalNumberTraces  = SUM_Trace_Recorded_OutputLog
            TotalDeadTraces    = SUM_Trace_Dead_OutputLog    
            YieldTracesPercent = Percent_Yield_OutputLog
            DeadTracepercent   = Percent_Dead_OutputLog

            Trace_Summary = pd.DataFrame({'Unique Record (FFID)':[TotalUniqueFFID],
                                          'Total Traces Recorded (#)':[TotalNumberTraces], 
                                          'Total Missing/Dead Traces (#)':[TotalDeadTraces],
                                          'Trace Recovery (Yield %)':[YieldTracesPercent],
                                          'Missing/Dead Traces (%)':[DeadTracepercent]
                                                   },index=None)

            def get_TraceYield_Summary_Rep_datetime():
                return "Geospace OutputLog Trace Yield Report -" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"

            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",
                                         title = "Save Geospace OutputLog Trace Yield Report As Excel",
                                         filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
            if len(filename) >0:
                TraceYieldSummary   = get_TraceYield_Summary_Rep_datetime()
                outfile_TraceYieldSummary  = filename + TraceYieldSummary
                XLSX_writer = pd.ExcelWriter(outfile_TraceYieldSummary)
                Trace_Summary.to_excel(XLSX_writer,'TraceYieldReport', startrow = 1 ,  index=False)
                Valid_No_Duplicated_OutputLog_DF.to_excel(XLSX_writer,'OutputLogReport', startrow = 0 ,  index=False)

                workbook       = XLSX_writer.book
                worksheet_Trace_Summary  = XLSX_writer.sheets['TraceYieldReport']
                worksheet_OutputLog      = XLSX_writer.sheets['OutputLogReport']   
                header1  = '&L&G'+'&CEagle Canada Seismic Services ULC' + '\n' + '6806 Railway Street SE' + '\n' + 'Calgary, AB T2H 3A8' + '\n' + 'Ph: (403) 263-7770' +  '&R&U&14&"cambria, bold"Trace Yield Report' + '\n'+  'Date : &D'
                header2  = '&L&G'+'&CEagle Canada Seismic Services ULC' + '\n' + '6806 Railway Street SE' + '\n' + 'Calgary, AB T2H 3A8' + '\n' + 'Ph: (403) 263-7770' +  '&R&U&14&"cambria, bold"Output Log Report' + '\n'+  'Date : &D'
                worksheet_Trace_Summary.set_header(header1,{'image_left':'eagle logo.jpg'})
                worksheet_OutputLog.set_header(header2,{'image_left':'eagle logo.jpg'})
                footer1  = ('&LDate : &D')        
                worksheet_Trace_Summary.set_footer(footer1)
                worksheet_OutputLog.set_footer(footer1)
                
                worksheet_Trace_Summary.set_landscape()
                worksheet_Trace_Summary.set_margins(0.6, 0.6, 1.6, 1.1)
                worksheet_Trace_Summary.print_area('A1:E30')
                worksheet_Trace_Summary.set_paper(9)
                worksheet_Trace_Summary.set_start_page(1)
                worksheet_Trace_Summary.hide_gridlines(1)
                worksheet_Trace_Summary.set_page_view()
                workbook.formats[0].set_align('center')
                workbook.formats[0].set_font_size(11)
                workbook.formats[0].set_bold(True)
                workbook.formats[0].set_border(1)
                worksheet_Trace_Summary.set_column(0, 0, 28)
                worksheet_Trace_Summary.set_column(1, 1, 28)
                worksheet_Trace_Summary.set_column(2, 2, 28)
                worksheet_Trace_Summary.set_column(3, 3, 28)
                worksheet_Trace_Summary.set_column(4, 4, 28)

                worksheet_OutputLog.set_landscape()
                worksheet_OutputLog.set_margins(0.6, 0.6, 1.6, 1.1)                                   
                worksheet_OutputLog.set_paper(9)
                worksheet_OutputLog.set_start_page(1)
                worksheet_OutputLog.set_page_view()
                workbook.formats[0].set_align('center')
                workbook.formats[0].set_font_size(10)
                workbook.formats[0].set_bold(True)
                workbook.formats[0].set_border(1)
                worksheet_OutputLog.set_column(0, 0, 11)
                worksheet_OutputLog.set_column(1, 2, 13)
                worksheet_OutputLog.set_column(3, 4, 11)
                worksheet_OutputLog.set_column(5, 5, 10)
                worksheet_OutputLog.set_column(6, 9, 17)
                
                
                cell_format_Left = workbook.add_format({
                                                        'bold': True,
                                                        'text_wrap': False,
                                                        'valign': 'top',
                                                        'border': 1})
                cell_format_Left.set_align('left')
                cell_format_Left.set_font_size(11)
                cell_format_Center = workbook.add_format({
                                                        'bold': True,
                                                        'text_wrap': False,
                                                        'valign': 'top',
                                                        'border': 1})
                cell_format_Center.set_align('center')
                cell_format_Center.set_font_size(11)
                cell_format_Header = workbook.add_format({
                                                        'bold': True,
                                                        'text_wrap': False,
                                                        'valign': 'top',
                                                        'border': 1})
                cell_format_Header.set_align('left')
                cell_format_Header.set_font_size(12)
                cell_format_Header.set_underline(1)
                cell_format_Footnote = workbook.add_format({
                                                        'bold': True,
                                                        'text_wrap': False,
                                                        'valign': 'top',
                                                        'border': 1})
                cell_format_Footnote.set_align('left')
                cell_format_Footnote.set_font_size(12)
                worksheet_Trace_Summary.merge_range('A1:E1', " Output Log Trace Yield Summary :", cell_format_Header)


                #### Plot ## of Traces Dead VS FFID
                traceYieldNumber_chart_Bar = workbook.add_chart({'type': 'line'})
                traceYieldNumber_chart_Bar.add_series({  
                            'categories': ['OutputLogReport', 1, 0, TotalUniqueFFID, 0],    
                            'values':     ['OutputLogReport', 1, 7, TotalUniqueFFID, 7], 'gap': 500})
                traceYieldNumber_chart_Bar.set_size({'width': 1000, 'height': 250})                                  
                traceYieldNumber_chart_Bar.set_y_axis({
                    'name': '# Of Missing/Dead Channels',
                    'name_font': {'size': 8, 'bold': True},
                    'num_font':  {'size': 8, 'bold': True, 'rotation': - 45},
                    'major_gridlines': {
                                  'visible': True,
                                  'line': {'width': 0.1, 'dash_type': 'dash'}},})                        
                traceYieldNumber_chart_Bar.set_x_axis({'name': 'File Number (FFID) Assending', 'num_font':  {'size': 8, 'bold': True, 'rotation': - 45}, 'name_font': {'size': 9, 'bold': True},
                                  'major_gridlines': {
                                  'visible': True,
                                  'line': {'width': 0.1, 'dash_type': 'dash'}},})
                traceYieldNumber_chart_Bar.set_style(10)
                traceYieldNumber_chart_Bar.set_legend({'none': True}) 
                worksheet_Trace_Summary.insert_chart('A4', traceYieldNumber_chart_Bar,  
                                {'x_offset': 1, 'y_offset': 10})

                 #### Plot %% of Traces Dead VS FFID
                traceYieldPercent_chart_Bar = workbook.add_chart({'type': 'line'})
                traceYieldPercent_chart_Bar.add_series({  
                            'categories': ['OutputLogReport', 1, 0, TotalUniqueFFID, 0],    
                            'values':     ['OutputLogReport', 1, 9, TotalUniqueFFID, 9], 'gap': 500})
                traceYieldPercent_chart_Bar.set_size({'width': 1000, 'height': 250})                                  
                traceYieldPercent_chart_Bar.set_y_axis({
                    'name': '% Of Missing/Dead Channels',
                    'name_font': {'size': 8, 'bold': True},
                    'num_font':  {'size': 8, 'bold': True, 'rotation': - 45},
                    'major_gridlines': {
                                  'visible': True,
                                  'line': {'width': 0.1, 'dash_type': 'dash'}},})                        
                traceYieldPercent_chart_Bar.set_x_axis({'name': 'File Number (FFID) Assending', 'num_font':  {'size': 8, 'bold': True, 'rotation': - 45}, 'name_font': {'size': 9, 'bold': True},
                                  'major_gridlines': {
                                  'visible': True,
                                  'line': {'width': 0.1, 'dash_type': 'dash'}},})
                traceYieldPercent_chart_Bar.set_style(10)
                traceYieldPercent_chart_Bar.set_legend({'none': True}) 
                worksheet_Trace_Summary.insert_chart('A17', traceYieldPercent_chart_Bar,  
                                {'x_offset': 1, 'y_offset': 10})

                XLSX_writer.save()
                XLSX_writer.close()
                tkinter.messagebox.showinfo("Geospace Trace Yield Report Export Message"," Geospace Trace Yield Report Saved As Excel")
            else:
                tkinter.messagebox.showinfo("Geospace Trace Yield Report Export Message","Please Select Geospace Trace Yield Report File Name To Save")

    def ViewTotalImport():    
        conn = sqlite3.connect("GeoSpaceOutputLogTraceYield.db")
        Complete_df = pd.read_sql_query("SELECT * FROM GeoSpaceOutputLog ORDER BY `FileNumber` ASC ;", conn)
        ListBox_DF = pd.DataFrame(Complete_df)
        TotalListBox= len(ListBox_DF)
        if TotalListBox >0:
            tree.delete(*tree.get_children())
            txtTotalUniqueFFID.delete(0,END)
            txtTotalSeismicTrace.delete(0,END)
            txtTotalVoidTraces.delete(0,END)
            txtTotalPercentRecovery.delete(0,END)
            txtTotalPercentVoid.delete(0,END)

            Valid_No_Duplicated_OutputLog_DF = pd.DataFrame(Complete_df)
            Valid_No_Duplicated_OutputLog_DF = Valid_No_Duplicated_OutputLog_DF.reset_index(drop=True)

            Length_No_Duplicated_OutputLog  = len(Valid_No_Duplicated_OutputLog_DF)
            SUM_Trace_Recorded_OutputLog    = ((Valid_No_Duplicated_OutputLog_DF['NumSeisChannels']).sum())
            SUM_Trace_Dead_OutputLog        = ((Valid_No_Duplicated_OutputLog_DF['NumMissChannels']).sum())
            Percent_Dead_OutputLog    = round((((SUM_Trace_Dead_OutputLog)/(SUM_Trace_Recorded_OutputLog))*100),2)
            Percent_Yield_OutputLog   = round((100 - Percent_Dead_OutputLog),2)            
            txtTotalUniqueFFID.insert(tk.END,Length_No_Duplicated_OutputLog)
            txtTotalSeismicTrace.insert(tk.END,SUM_Trace_Recorded_OutputLog)
            txtTotalVoidTraces.insert(tk.END,SUM_Trace_Dead_OutputLog)
            txtTotalPercentVoid.insert(tk.END,Percent_Dead_OutputLog)
            txtTotalPercentRecovery.insert(tk.END,Percent_Yield_OutputLog)            
            Valid_No_Duplicated_OutputLog_DF = pd.DataFrame(Valid_No_Duplicated_OutputLog_DF)
            Valid_No_Duplicated_OutputLog_DF  = Valid_No_Duplicated_OutputLog_DF.sort_values(by =['FileNumber'])
            Valid_No_Duplicated_OutputLog_DF = Valid_No_Duplicated_OutputLog_DF.reset_index(drop=True)
            for each_rec in range(len(Valid_No_Duplicated_OutputLog_DF)):
                tree.insert("", tk.END, values=list(Valid_No_Duplicated_OutputLog_DF.loc[each_rec]))
        conn.commit()
        conn.close()

    def ViewDuplicatedImport():
        tree.delete(*tree.get_children())
        txtDuplictedTraces.delete(0,END)    
        conn = sqlite3.connect("GeoSpaceOutputLogTraceYield.db")
        Complete_df = pd.read_sql_query("SELECT * FROM GeoSpaceOutputLog_Duplicated ORDER BY `FileNumber` ASC ;", conn)
        Valid_Duplicated_OutputLog_DF = pd.DataFrame(Complete_df)   
        Length_Duplicated_OutputLog = len(Valid_Duplicated_OutputLog_DF)    
        txtDuplictedTraces.insert(tk.END,Length_Duplicated_OutputLog)
        if Length_Duplicated_OutputLog >0:        
            Valid_Duplicated_OutputLog_DF  = Valid_Duplicated_OutputLog_DF.sort_values(by =['FileNumber'])
            Valid_Duplicated_OutputLog_DF  = Valid_Duplicated_OutputLog_DF.reset_index(drop=True)
            for each_rec in range(len(Valid_Duplicated_OutputLog_DF)):
                tree.insert("", tk.END, values=list(Valid_Duplicated_OutputLog_DF.loc[each_rec]))

        conn.commit()
        conn.close()    
               
    def SortbyLineStation():
        dfList =[] 
        for child in tree.get_children():
            df = tree.item(child)["values"]
            dfList.append(df)
        ListBox_DF = pd.DataFrame(dfList)
        ListBox_DF.rename(columns={0:'FileNumber',     1:'ShotOverrides',  2:'ProcessType',    3:'ShotLine', 4:'ShotStation', 5:'Comment',
                                   6:'NumSeisChannels',7:'NumMissChannels',8:'NumZeroChannels',9:'PercentMissing'},inplace = True)
                    
        SortbyLineStation_ListBox  = pd.DataFrame(ListBox_DF)    
        TotalListBox= len(SortbyLineStation_ListBox)       
        if TotalListBox >0:
            SortbyLineStation_ListBox  = SortbyLineStation_ListBox.sort_values(by =['ShotLine', 'ShotStation'], ascending=True)
            SortbyLineStation_ListBox  = SortbyLineStation_ListBox.reset_index(drop=True)
            tree.delete(*tree.get_children())
            for each_rec in range(len(SortbyLineStation_ListBox)):
                    tree.insert("", tk.END, values=list(SortbyLineStation_ListBox.loc[each_rec]))
    def SortbyPercentMissing():
        dfList =[] 
        for child in tree.get_children():
            df = tree.item(child)["values"]
            dfList.append(df)
        ListBox_DF = pd.DataFrame(dfList)
        ListBox_DF.rename(columns={0:'FileNumber',     1:'ShotOverrides',  2:'ProcessType',    3:'ShotLine',    4:'ShotStation', 5:'Comment',
                                   6:'NumSeisChannels',7:'NumMissChannels',8:'NumZeroChannels',9:'PercentMissing'},inplace = True)
                    
        SortbyLineStation_ListBox  = pd.DataFrame(ListBox_DF)    
        TotalListBox= len(SortbyLineStation_ListBox)       
        if TotalListBox >0:
            SortbyLineStation_ListBox  = SortbyLineStation_ListBox.sort_values(by =['NumMissChannels', 'PercentMissing'], ascending=False)
            SortbyLineStation_ListBox  = SortbyLineStation_ListBox.reset_index(drop=True)
            tree.delete(*tree.get_children())
            for each_rec in range(len(SortbyLineStation_ListBox)):
                    tree.insert("", tk.END, values=list(SortbyLineStation_ListBox.loc[each_rec]))
    def PlotTraceReport():
        dfList =[] 
        for child in tree.get_children():
            df = tree.item(child)["values"]
            dfList.append(df)
        ListBox_DF = pd.DataFrame(dfList)
        TotalListBox= len(ListBox_DF)
        if TotalListBox >0:
            ListBox_DF.rename(columns={0:'FileNumber', 1:'ShotOverrides', 2:'ProcessType', 3:'ShotLine', 4:'ShotStation', 5:'Comment',
                                     6:'NumSeisChannels',7:'Number Of Missing Channels (#)',8:'Number Of Zero Channels (#)',9:'Percent Of Missing/Dead Channels (%)'},inplace = True)
            YieldTracesPercent = txtTotalPercentRecovery.get()
            DeadTracepercent   = txtTotalPercentVoid.get()                
            PlotZeroChannelsVSFFID  = pd.DataFrame(ListBox_DF)
            PlotZeroChannelsVSFFID  = PlotZeroChannelsVSFFID.sort_values(by =['FileNumber'], ascending=True)
            PlotZeroChannelsVSFFID['Number Of Missing/Dead Channels (#)']   = (PlotZeroChannelsVSFFID.loc[:,['Number Of Missing Channels (#)']]).astype(int)
            PlotZeroChannelsVSFFID['Percent Of Missing/Dead Channels (%)']  = (PlotZeroChannelsVSFFID.loc[:,['Percent Of Missing/Dead Channels (%)']]).astype(float)
            Plot_XY = PlotZeroChannelsVSFFID.plot(x='FileNumber', y=['Number Of Missing/Dead Channels (#)', 'Percent Of Missing/Dead Channels (%)'] ,layout=None,  subplots=True, color='red',
                                        title=("Trace Recovery (Yield %) : " + YieldTracesPercent + '    ' + " Trace Missing/Dead (%) : " + DeadTracepercent),
                                        legend=True, ylim=None, fontsize=6)
            plt.legend(loc='best')
            plt.xlabel("File Number (FFID) Assending", fontsize=8)
            Plot_XY[0].set_ylabel("# Of Missing/Dead Channels", fontsize=8)
            Plot_XY[1].set_ylabel("% Of Missing/Dead Channels", fontsize=8)   
           
            plt.rcParams.update({'figure.max_open_warning': 0})
            plt.grid(True)
            plt.show()
    def ImportOutputLogFile():
        ClearView()
        fileList = askopenfilenames(initialdir = "/", title = "Import Geospace Output Log Files" , filetypes=[('Excel File', ('*.xls', '*.xlsx')), ('CSV File', '*.csv')])
        Length_fileList  =  len(fileList)
        if Length_fileList >0:            
            if fileList:
                dfList =[]            
                for filename in fileList:
                    if filename.endswith('.csv'):
                        df = pd.read_csv(filename, sep=',' , low_memory=False)
                        df = df.iloc[:,:]
                        FileNumber          = df.loc[:,'File Number']
                        ShotOverrides       = df.loc[:,'Shot Overrides']
                        ProcessType         = df.loc[:,'Process Type']
                        ShotLine            = df.loc[:,'Line Number']
                        ShotStation         = df.loc[:,'Station Number']
                        Comment             = df.loc[:,'Comments']
                        NumSeisChannels     = df.loc[:,'Num Seis Channels']
                        NumMissChannels     = df.loc[:,'Num Miss Channels']
                        NumZeroChannels     = df.loc[:,'Num Zero Channels']
                        
                        column_names = [FileNumber, ShotOverrides,ProcessType, ShotLine, ShotStation, Comment, NumSeisChannels, NumMissChannels, NumZeroChannels]
                        catdf = pd.concat (column_names,axis=1,ignore_index =True)
                        dfList.append(catdf) 
                    else:
                        df = pd.read_excel(filename)
                        df = df.iloc[:,:]
                        FileNumber          = df.loc[:,'File Number']
                        ShotOverrides       = df.loc[:,'Shot Overrides']
                        ProcessType         = df.loc[:,'Process Type']
                        ShotLine            = df.loc[:,'Line Number']
                        ShotStation         = df.loc[:,'Station Number']
                        Comment             = df.loc[:,'Comments']
                        NumSeisChannels     = df.loc[:,'Num Seis Channels']
                        NumMissChannels     = df.loc[:,'Num Miss Channels']
                        NumZeroChannels     = df.loc[:,'Num Zero Channels']
                        column_names = [FileNumber, ShotOverrides,ProcessType, ShotLine, ShotStation, Comment, NumSeisChannels, NumMissChannels, NumZeroChannels]
                        catdf = pd.concat (column_names,axis=1,ignore_index =True)
                        dfList.append(catdf) 

                concatDf = pd.concat(dfList,axis=0, ignore_index =True)
                concatDf.rename(columns={0:'FileNumber', 1:'ShotOverrides', 2:'ProcessType', 3:'ShotLine', 4:'ShotStation', 5:'Comment',
                                 6:'NumSeisChannels',7:'NumMissChannels',8:'NumZeroChannels'},inplace = True)
                # RAW DUMP Total PSS imported
                RAW_DUMP_ImportedOutputLog_DF    = pd.DataFrame(concatDf)

                # Separating Valid with FFID No Null and No Duplicated FFID
                Valid_OutputLog_DF = pd.DataFrame(RAW_DUMP_ImportedOutputLog_DF)
                Valid_OutputLog_DF = Valid_OutputLog_DF[pd.to_numeric(Valid_OutputLog_DF.FileNumber, errors='coerce').notnull()]
                Valid_OutputLog_DF['DuplicatedEntries']      = Valid_OutputLog_DF.sort_values(by =['FileNumber']).duplicated(['FileNumber','ShotLine','ShotStation','ProcessType', 'ShotOverrides'],keep='last')
                Valid_OutputLog_DF                           = Valid_OutputLog_DF.reset_index(drop=True)
                Valid_OutputLog_DF                           = pd.DataFrame(Valid_OutputLog_DF)

                # Separated Valid with FFID No Null and Duplicated FFID
                Valid_Duplicated_OutputLog_DF = Valid_OutputLog_DF.loc[Valid_OutputLog_DF.DuplicatedEntries == True, 'FileNumber': 'NumZeroChannels']
                Valid_Duplicated_OutputLog_DF = Valid_Duplicated_OutputLog_DF.reset_index(drop=True)
                Valid_Duplicated_OutputLog_DF = pd.DataFrame(Valid_Duplicated_OutputLog_DF)
                Valid_Duplicated_OutputLog_DF['PercentMissing'] = (((Valid_Duplicated_OutputLog_DF['NumMissChannels'])/(Valid_Duplicated_OutputLog_DF['NumSeisChannels']))*100).round(2)

                # Separated Valid with FFID No Null and No Duplicated FFID
                Valid_No_Duplicated_OutputLog_DF = Valid_OutputLog_DF.loc[Valid_OutputLog_DF.DuplicatedEntries == False, 'FileNumber': 'NumZeroChannels']
                Valid_No_Duplicated_OutputLog_DF = Valid_No_Duplicated_OutputLog_DF.reset_index(drop=True)
                Valid_No_Duplicated_OutputLog_DF = pd.DataFrame(Valid_No_Duplicated_OutputLog_DF)
                Valid_No_Duplicated_OutputLog_DF['PercentMissing'] = (((Valid_No_Duplicated_OutputLog_DF['NumMissChannels'])/(Valid_No_Duplicated_OutputLog_DF['NumSeisChannels']))*100).round(2)

                ## TreeView Populated
                Length_No_Duplicated_OutputLog = len(Valid_No_Duplicated_OutputLog_DF)
                Length_Duplicated_OutputLog    = len(Valid_Duplicated_OutputLog_DF)
                SUM_Trace_Recorded_OutputLog   = ((Valid_No_Duplicated_OutputLog_DF['NumSeisChannels']).sum())
                SUM_Trace_Dead_OutputLog       = ((Valid_No_Duplicated_OutputLog_DF['NumMissChannels']).sum())
                Percent_Dead_OutputLog         = round((((SUM_Trace_Dead_OutputLog)/(SUM_Trace_Recorded_OutputLog))*100),2)
                Percent_Yield_OutputLog        = round((100 -Percent_Dead_OutputLog),2)            
                txtTotalUniqueFFID.insert(tk.END,Length_No_Duplicated_OutputLog)
                txtDuplictedTraces.insert(tk.END,Length_Duplicated_OutputLog)
                txtTotalSeismicTrace.insert(tk.END,SUM_Trace_Recorded_OutputLog)
                txtTotalVoidTraces.insert(tk.END,SUM_Trace_Dead_OutputLog)
                txtTotalPercentVoid.insert(tk.END,Percent_Dead_OutputLog)
                txtTotalPercentRecovery.insert(tk.END,Percent_Yield_OutputLog)            
                Valid_No_Duplicated_OutputLog_DF = pd.DataFrame(Valid_No_Duplicated_OutputLog_DF)
                Valid_No_Duplicated_OutputLog_DF  = Valid_No_Duplicated_OutputLog_DF.sort_values(by =['FileNumber'])
                Valid_No_Duplicated_OutputLog_DF = Valid_No_Duplicated_OutputLog_DF.reset_index(drop=True)
                for each_rec in range(len(Valid_No_Duplicated_OutputLog_DF)):
                    tree.insert("", tk.END, values=list(Valid_No_Duplicated_OutputLog_DF.loc[each_rec]))

                # Connect To Database and Export DF  
                con= sqlite3.connect("GeoSpaceOutputLogTraceYield.db")
                cur=con.cursor()
                Valid_No_Duplicated_OutputLog_DF.to_sql('GeoSpaceOutputLog',con, if_exists="replace", index=False)
                Valid_Duplicated_OutputLog_DF.to_sql('GeoSpaceOutputLog_Duplicated',con, if_exists="replace", index=False)
                
                con.commit()
                cur.close()
                con.close()
        else:
            tkinter.messagebox.showinfo("Import Output Log File Message","Please Select Output Log Files To Import")
    def ClearView():
        tree.delete(*tree.get_children())
        txtTotalUniqueFFID.delete(0,END)
        txtDuplictedTraces.delete(0,END)
        txtTotalSeismicTrace.delete(0,END)
        txtTotalVoidTraces.delete(0,END)
        txtTotalPercentRecovery.delete(0,END)
        txtTotalPercentVoid.delete(0,END)

    ## DataFrame TOP
    Label_txtUniqueFFID = Label(DataFrameTOP, text = "Total Unique Record :", font=("arial", 10,'bold'), bg = 'cadet blue')
    Label_txtUniqueFFID.grid(row =0, column = 0, sticky ="W", padx= 1, pady =5)
    txtTotalUniqueFFID   = Entry(DataFrameTOP, font=('aerial', 10, 'bold'),textvariable=IntVar(), width = 10, bd=2)
    txtTotalUniqueFFID.grid(row =0, column = 0, sticky ="W", padx= 150 , pady =5)
    btnViewUniqueFFID = Button(DataFrameTOP, text="View", font=('aerial', 9, 'bold'), height =1, width=4, bd=1, command = ViewTotalImport)
    btnViewUniqueFFID.grid(row =0, column = 0, sticky ="W", padx= 230 , pady =5)

    Label_txtDuplictedTraces = Label(DataFrameTOP, text = "Duplicated Record:", font=("arial", 10,'bold'), bg = 'cadet blue')
    Label_txtDuplictedTraces.grid(row =2, column = 0, sticky ="W", padx= 1, pady =5)
    txtDuplictedTraces  = Entry(DataFrameTOP, font=('aerial', 10, 'bold'),textvariable=IntVar(), width = 10, bd=2)
    txtDuplictedTraces.grid(row =2, column = 0, sticky ="W", padx= 150 , pady =5)
    btnViewDuplicatedFFID = Button(DataFrameTOP, text="View", font=('aerial', 9, 'bold'), height =1, width=4, bd=1, command = ViewDuplicatedImport)
    btnViewDuplicatedFFID.grid(row =2, column = 0, sticky ="W", padx= 230 , pady =5)

    Label_txtTotalSeismicTrace = Label(DataFrameTOP, text = "Traces Recorded :", font=("arial", 10,'bold'), bg = 'cadet blue')
    Label_txtTotalSeismicTrace.grid(row =0, column = 4, sticky ="W", padx= 1, pady =5)
    txtTotalSeismicTrace  = Entry(DataFrameTOP, font=('aerial', 10, 'bold'),textvariable=IntVar(), width = 15, bd=2)
    txtTotalSeismicTrace.grid(row =0, column = 4, sticky ="W", padx= 145 , pady =5)

    Label_txtVoidTraces = Label(DataFrameTOP, text = "Missing/Dead Traces:", font=("arial", 10,'bold'), bg = 'cadet blue')
    Label_txtVoidTraces.grid(row =2, column = 4, sticky ="W", padx= 1, pady =5)
    txtTotalVoidTraces  = Entry(DataFrameTOP, font=('aerial', 10, 'bold'),textvariable=IntVar(), width = 15, bd=2)
    txtTotalVoidTraces.grid(row =2, column = 4, sticky ="W", padx= 145 , pady =5)

    Label_txtPercentRecovery = Label(DataFrameTOP, text = " Trace Recovery Percent (Yield %) :", font=("arial", 10,'bold'), bg = 'cadet blue')
    Label_txtPercentRecovery.grid(row =0, column = 8, sticky ="W", padx= 18, pady =5)
    txtTotalPercentRecovery   = Entry(DataFrameTOP, font=('aerial', 10, 'bold'),textvariable=IntVar(), width = 9, bd=2)
    txtTotalPercentRecovery.grid(row =0, column = 8, sticky ="W", padx= 255 , pady =5)

    Label_txtPercentVoid = Label(DataFrameTOP, text = "Missing/Dead Traces Percent (%) :", font=("arial", 10,'bold'), bg = 'cadet blue')
    Label_txtPercentVoid.grid(row =2, column = 8, sticky ="W", padx= 20, pady =5)
    txtTotalPercentVoid   = Entry(DataFrameTOP, font=('aerial', 10, 'bold'),textvariable=IntVar(), width = 9, bd=2)
    txtTotalPercentVoid.grid(row =2, column = 8, sticky ="W", padx= 255 , pady =5)

    DataFrameTOP.pack()

    ## DataFrame BOTTOM ACTIONS
    btnImportOutputLog = Button(DataFrameBOTTOM_ACTIONS, text="Import Output Log", font=('aerial', 10, 'bold'), height =1, width=16, bd=2, command = ImportOutputLogFile)
    btnImportOutputLog.grid(row =0, column = 0, sticky ="W", padx= 1 , pady =0)

    btnExportTraceReport = Button(DataFrameBOTTOM_ACTIONS, text="Export Trace Report", font=('aerial', 10, 'bold'), height =1, width=16, bd=2, command = ExportTraceReport)
    btnExportTraceReport.grid(row =0, column = 1, sticky ="W", padx= 1)

    btnPlotTraceReport = Button(DataFrameBOTTOM_ACTIONS, text="Plot Trace Report", font=('aerial', 10, 'bold'), height =1, width=16, bd=2, command = PlotTraceReport)
    btnPlotTraceReport.grid(row =0, column = 2, sticky ="W", padx= 1)

    btnSortbyLineStation = Button(DataFrameBOTTOM_ACTIONS, text="Sort By Shot Point", font=('aerial', 10, 'bold'), height =1, width=16, bd=2, command = SortbyLineStation)
    btnSortbyLineStation.grid(row =0, column = 3, sticky ="W", padx= 1)

    btnSortbyPercentMissing = Button(DataFrameBOTTOM_ACTIONS, text="Sort By Missing Traces", font=('aerial', 10, 'bold'), height =1, width=20, bd=2, command = SortbyPercentMissing)
    btnSortbyPercentMissing.grid(row =0, column = 4, sticky ="W", padx= 1)

    DataFrameBOTTOM_ACTIONS.pack()






