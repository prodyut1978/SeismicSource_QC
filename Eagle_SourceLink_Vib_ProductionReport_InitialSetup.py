import os
import sys
import PySimpleGUI as sg
import pickle
from tkinter import*
import tkinter.messagebox
import pandas as pd
import glob
import datetime
from datetime import datetime
import csv
import openpyxl
import numpy as np
import tkinter as tk
import sqlite3
import tkinter.ttk as ttk
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilenames
from tkinter import simpledialog
from PyPDF2 import PdfFileMerger
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.backends.backend_pdf

def VibProductionReportInitialSetup():  
    Vib_VibProductionReport_path = r'C:\VibProductionReport'
    if not os.path.exists(Vib_VibProductionReport_path):
        os.makedirs(Vib_VibProductionReport_path)

    Vib_VibProductionReport_path = r'C:\VibRestrictedFolder\VibProductionModule'
    if not os.path.exists(Vib_VibProductionReport_path):
        os.makedirs(Vib_VibProductionReport_path)

    Vib_VibProductionReport_path_Plot_File = r'C:\VibProductionReport\PlotFiles'
    if not os.path.exists(Vib_VibProductionReport_path_Plot_File):
        os.makedirs(Vib_VibProductionReport_path_Plot_File)
    PlotFolder = [ f for f in os.listdir(Vib_VibProductionReport_path_Plot_File) if f.endswith(".pdf") ]
    for f in PlotFolder:
        os.remove(os.path.join(Vib_VibProductionReport_path_Plot_File, f))

    VibTemp_path = r'C:\VibProductionReport\Support_Files\TempPlotFiles'
    if not os.path.exists(VibTemp_path):
        os.makedirs(VibTemp_path)
    DetailTempfilelist = [ f for f in os.listdir(VibTemp_path) if f.endswith(".pdf") ]
    for f in DetailTempfilelist:
        os.remove(os.path.join(VibTemp_path, f))

    Default_Date_today   = "Date : " + datetime.now().strftime("%Y-%m-%d")


    layout = [[sg.Text('Enter Total Number Of Vibes In Crew:',    size=(40,1)), sg.InputText()],        
                  [sg.Submit(), sg.Cancel()]]
    WindowRead = sg.Window('Input Number Of Vibes',
             auto_size_text=True, default_element_size=(10, 1)).Layout(layout)

    event, values = WindowRead.read()


    if event is None or event == 'Cancel':
        sg.PopupAutoClose('Exiting Vib Number Input',line_width=60)
        
    else:
        TotalNumberVibes = values[0]
        if (len(TotalNumberVibes) == 0):
            tkinter.messagebox.showinfo("TotalNumberVibes Error Message","TotalNumberVibes Can Not Be Empty")
        else:
            try:
                TotalNumberVibes = int(values[0])
                TotalNumberVibes = TotalNumberVibes + 1
                try:
                    layout = []
                    for i in range(1,TotalNumberVibes):
                        layout += [sg.Text(f'{i}. ' "Enter Vib" f'{i}. ' "PSS Profile ID :" , size=(25,1)),  sg.In(i),
                                   sg.Text(f'      {i}. ' "Enter Vib" f'{i}. ' "Unit No :" , size=(22,1)),  sg.In()],
                    layout += [[sg.Button('Submit'), sg.Button('Cancel')]]

                    ReadwindowVibInfo   = sg.Window('Vib Input Profile', auto_size_text=True, default_element_size=(10, 1)).Layout(layout)
                    event, values       = ReadwindowVibInfo.read()
                    
                    VibProfileParameter = values
                    VibProfileParameter = { i : VibProfileParameter[i] for i in range(0, len(VibProfileParameter) ) }
                    pickle_out          = open("C:\VibRestrictedFolder\VibProductionModule\VibProfileParameter","wb")        
                    pickle.dump(VibProfileParameter,pickle_out)
                    pickle_out.close()                

                    if event is None or event == 'Cancel':
                            sg.PopupAutoClose('Exiting Vib Input QC',line_width=60)
                            
                    else:
                        Vib_ProfileID    = []
                        VibNumberList    = ["Vib No : → "]
                        VibNumberPlotDF  = []
                        Len_Values       = len(values)

                        for i in range(0, Len_Values, 2):
                            Vib_ProfileID.append(values[i])
                            
                        for i in range(1, Len_Values, 2):
                            VibNumberList.append(values[i])
                            VibNumberPlotDF.append(values[i]) 
                            
                        Number_of_Vib        = pd.DataFrame({'Unit_ID': Vib_ProfileID})
                        Number_of_Vib        = Number_of_Vib.replace('','Empty', regex=True)

                        Vib_Number           = pd.DataFrame({'Vib_Number': VibNumberList})
                        Vib_Number           = Vib_Number.replace('','Empty', regex=True)

                        Vib_NumberPlot       = pd.DataFrame({'Vib_Number': VibNumberPlotDF})
                        Vib_NumberPlot       = Vib_NumberPlot.replace('','Empty', regex=True)

                        Empty_Check_Profile  = Number_of_Vib['Unit_ID']== 'Empty'
                        Empty_Check_VibNum   = Vib_Number['Vib_Number']== 'Empty'
                        Empty_Check_VibPlot  = Vib_NumberPlot['Vib_Number']== 'Empty'

                        if (Empty_Check_Profile.values.any()==True) | (Empty_Check_VibNum.values.any()==True) | (Empty_Check_VibPlot.values.any()==True) | (Vib_NumberPlot['Vib_Number'].duplicated().values.any() == True) | (Number_of_Vib['Unit_ID'].duplicated().values.any() == True):
                            tkinter.messagebox.showinfo("Input File Error Message","Please Input All Vib Profile ID And vib Number Correctly : Duplicated Entry or Empty Values")

                        else:                    
                            Number_of_Vib                  = pd.DataFrame(Number_of_Vib)
                            Number_of_Vib                  = Number_of_Vib.reset_index(drop=True)
                            Number_of_Vib['Unit_ID']       = Number_of_Vib['Unit_ID'].astype(int)
                            Vib_Number                     = Vib_Number.T    
                            Vib_NumberPlotly               = pd.DataFrame(Vib_NumberPlot)
                            Vib_NumberPlotly               = Vib_NumberPlotly.reset_index(drop=True)
                            Vib_NumberPlotly['Vib_Number'] = Vib_NumberPlotly['Vib_Number'].astype(int)
                            fileList_PSS = askopenfilenames(initialdir = "/", title = "Import PSS Files" , filetypes=[('PSS CSV File', '*.csv')])
                            Length_fileList  =  len(fileList_PSS)
                            if Length_fileList >0:                             
                                dfList_PSS   = []
                                for filenamePSS in fileList_PSS:
                                    df                        = pd.read_csv(filenamePSS, sep=',', low_memory=False)
                                    df                        = df.iloc[:,:]    
                                    Number_of_Shots           = df.loc[:,'Shot ID']    
                                    PSS_Local_Date            = df.loc[:,'Date']    
                                    Unit_ID                   = df.loc[:,'Unit ID']
                                    Identifier                = PSS_Local_Date[0]
                                    ProductionDay             = Number_of_Shots.shape[0]*[Identifier]
                                    Production_Day            = pd.DataFrame(ProductionDay)        
                                    column_names_PSS          = [Number_of_Shots, PSS_Local_Date, Unit_ID, Production_Day]
                                    catdfPSS                  = pd.concat (column_names_PSS, axis=1, ignore_index =True)
                                    dfList_PSS.append(catdfPSS)        
                                concatDfPSS = pd.concat(dfList_PSS,axis=0)
                                concatDfPSS.rename(columns={0:'Number_of_Shots', 1:'PSS_Local_Date',
                                                            2:'Unit_ID', 3:'Production_Day'},inplace = True)
                                VIB_PSS = pd.DataFrame(concatDfPSS)
                                VIB_PSS = VIB_PSS[pd.to_numeric(VIB_PSS.Number_of_Shots,errors='coerce').notnull()]
                                VIB_PSS['DuplicatedEntries'] = VIB_PSS.duplicated(['Number_of_Shots'],keep='last')
                                VIB_PSS = VIB_PSS.loc[VIB_PSS.DuplicatedEntries == False, 'Number_of_Shots': 'Production_Day']
                                VIB_PSS = VIB_PSS.reset_index(drop=True)
                                VIB_PSS['Production_Day'] = pd.to_datetime(VIB_PSS.Production_Day)
                                VIB_PSS['Production_Day'] = VIB_PSS['Production_Day'].dt.strftime('%Y/%m/%d')
                                VIB_PSS                   = VIB_PSS.reset_index(drop=True)
                                VIB_PSS                   = pd.DataFrame(VIB_PSS)
                                Number_Shots_Per_vib    = VIB_PSS.groupby(['Unit_ID', 'Production_Day']).Number_of_Shots.count()
                                Number_Shots_Per_vib    = Number_Shots_Per_vib.reset_index(drop=False)
                                Number_Shots_Per_vib    = Number_Shots_Per_vib.loc[:,['Unit_ID', 'Production_Day', 'Number_of_Shots']]
                                Number_Shots_Per_vib    = pd.DataFrame(Number_Shots_Per_vib, index =None)
                                Number_Shots_Per_vib    = pd.merge(Number_of_Vib, Number_Shots_Per_vib, how='left', on='Unit_ID')
                                Number_Shots_Per_vib    = Number_Shots_Per_vib.reset_index(drop=True)
                                Number_Shots_Per_vib    = pd.DataFrame(Number_Shots_Per_vib)
                                Total_ProductionShots   = len(VIB_PSS)
                                Total_ProductionShots   = 'TotalShots : ' + str(Total_ProductionShots)
                                            
                                VibProfileIDList        = list((Number_Shots_Per_vib['Unit_ID']).unique())
                                LengthVibProfileIDList  = len (VibProfileIDList)
                                VibNumberList           = list((Vib_NumberPlot['Vib_Number']).unique())
                                LengthVibNumberList     = len (VibNumberList)
                                BarhInitial             = LengthVibProfileIDList+1

                                ## Bar Ploting Vib Production Report
                                for i in range(len(VibProfileIDList)):
                                    Per_vib  =  Number_Shots_Per_vib[(Number_Shots_Per_vib.Unit_ID == VibProfileIDList[i])]
                                    VibProfileID = VibProfileIDList[i]
                                    VibNumber    = VibNumberList [i]    
                                    df = Per_vib.set_index('Production_Day')[['Number_of_Shots']].plot.bar(rot=0)  
                                    plt.title(("Vib Profile ID: " + str(VibProfileID) + '      ' + "Vib Number: " + str(VibNumber)),size=14, verticalalignment='bottom')

                                    plt.xticks(fontsize=5, rotation = 70)
                                    plt.yticks(fontsize=5, rotation = 0)
                                    
                                    plt.xlabel('Production Per Day',fontsize=8)
                                    plt.ylabel('Number of Shots' ,fontsize=8)

                                    Length_Xticks = len(df.xaxis.get_ticklabels())
                                    every_nth     = 1 + round(Length_Xticks/10)

                                    for n, label in enumerate(df.xaxis.get_ticklabels()):
                                        if n % every_nth != 0:
                                            label.set_visible(False)
                    
                                    red_patch = mpatches.Patch(label='Shots Per Day')
                                    plt.legend(handles=[red_patch])
                                    plt.rcParams.update({'figure.max_open_warning': 0})
                                    def get_ProfileID_PDF():
                                            return "Bar Plot VibProfileID -" + str(VibProfileID) + ".pdf"
                                    Plt_get_ProfileID_PDF = get_ProfileID_PDF()
                                    Plt_path_PDF     = "C:\\VibProductionReport\\PlotFiles\\" + Plt_get_ProfileID_PDF
                                    TempPlt_path_PDF = "C:\\VibProductionReport\\Support_Files\\TempPlotFiles\\" + Plt_get_ProfileID_PDF
                                    plt.savefig((Plt_path_PDF), dpi=10, orientation='portrait',
                                                bbox_inches='tight',metadata={'Title': 'Vib Production Summary'})
                                    plt.savefig((TempPlt_path_PDF), dpi=10, orientation='portrait',
                                                bbox_inches='tight',metadata={'Title': 'Vib Production Summary'})

                                    def get_ProfileIDPNG():
                                            return "Bar Plot VibProfileID -" + str(VibProfileID) + ".png"
                                    Plt_get_ProfileIDPng = get_ProfileIDPNG()
                                    Plt_path_PNG     = "C:\\VibProductionReport\\PlotFiles\\" + Plt_get_ProfileIDPng                        
                                    plt.savefig((Plt_path_PNG), orientation='portrait' , bbox_inches='tight', metadata={'Title': 'Vib Production Summary'})

                                def get_VibProductionSummary_Plot_datetime():
                                    return "Vib Production Summary Bar Plot -" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".pdf"
                                VibProductionSummaryPlotCombined          = get_VibProductionSummary_Plot_datetime()        
                                outfile_VibProductionSummaryPlotCombined  = "C:\\VibProductionReport\\" + VibProductionSummaryPlotCombined  
                                pdf = matplotlib.backends.backend_pdf.PdfPages(outfile_VibProductionSummaryPlotCombined)
                                for fig in range(1, plt.gcf().number + 1):
                                    pdf.savefig( fig, dpi=10, orientation='portrait', bbox_inches='tight',metadata={'Title': 'Vib Production Summary'} )
                                pdf.close()


                                ## BarH Ploting Vib Production Report
                                for i in range(len(VibProfileIDList)):
                                    Per_vib  =  Number_Shots_Per_vib[(Number_Shots_Per_vib.Unit_ID == VibProfileIDList[i])]
                                    VibProfileID = VibProfileIDList[i]
                                    VibNumber    = VibNumberList [i]    
                                    df_BarH = Per_vib.set_index('Production_Day')[['Number_of_Shots']].plot.barh(rot=0)  
                                    plt.title(("Vib Profile ID: " + str(VibProfileID) + '      ' + "Vib Number: " + str(VibNumber)),size=14, verticalalignment='bottom')

                                    plt.xticks(fontsize=5, rotation=0)
                                    plt.yticks(fontsize=5, rotation=0) 

                                    plt.xlabel('Number of Shots' ,fontsize=8)
                                    plt.ylabel('Production Per Day',fontsize=8)

                                    Length_Yticks = len(df_BarH.yaxis.get_ticklabels())
                                    every_nth     = 1 + round(Length_Yticks/10)
                                    for n, label in enumerate(df_BarH.yaxis.get_ticklabels()):
                                        if n % every_nth != 0:
                                            label.set_visible(False)

                                    red_patch = mpatches.Patch(label='Shots Per Day')
                                    plt.legend(handles=[red_patch])
                                    plt.rcParams.update({'figure.max_open_warning': 0})
                                    def get_ProfileID_PDF():
                                            return "Barh Plot VibProfileID -" + str(VibProfileID) + ".pdf"
                                    Plt_get_ProfileID_PDF = get_ProfileID_PDF()
                                    Plt_path_PDF     = "C:\\VibProductionReport\\PlotFiles\\" + Plt_get_ProfileID_PDF
                                    TempPlt_path_PDF = "C:\\VibProductionReport\\Support_Files\\TempPlotFiles\\" + Plt_get_ProfileID_PDF
                                    plt.savefig((Plt_path_PDF), dpi=10, orientation='portrait',
                                                bbox_inches='tight',metadata={'Title': 'Vib Production Summary'})
                                    plt.savefig((TempPlt_path_PDF), dpi=10, orientation='portrait',
                                                bbox_inches='tight',metadata={'Title': 'Vib Production Summary'})

                                    def get_ProfileIDPNG():
                                            return "Barh Plot VibProfileID -" + str(VibProfileID) + ".png"
                                    Plt_get_ProfileIDPng = get_ProfileIDPNG()
                                    Plt_path_PNG     = "C:\\VibProductionReport\\PlotFiles\\" + Plt_get_ProfileIDPng                        
                                    plt.savefig((Plt_path_PNG), orientation='portrait' , bbox_inches='tight',metadata={'Title': 'Vib Production Summary'})

                                def get_VibProductionSummary_Plot_datetime():
                                    return "Vib Production Summary Barh Plot-" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".pdf"
                                VibProductionSummaryPlotCombined          = get_VibProductionSummary_Plot_datetime()        
                                outfile_VibProductionSummaryPlotCombined  = "C:\\VibProductionReport\\" + VibProductionSummaryPlotCombined  
                                pdf = matplotlib.backends.backend_pdf.PdfPages(outfile_VibProductionSummaryPlotCombined)
                                for fig in range(BarhInitial, plt.gcf().number + 1):
                                    pdf.savefig( fig, dpi=10, orientation='portrait', bbox_inches='tight',metadata={'Title': 'Vib Production Summary'} )
                                pdf.close()                    
                                

                                ## Generating Pivot Table With Vib Report
                                Number_Shots_Per_vib_Pivot    = pd.pivot_table(Number_Shots_Per_vib, values = 'Number_of_Shots', index = ['Production_Day'], columns = ["Unit_ID"] , aggfunc = np.sum,
                                                               fill_value = 0, dropna = False, margins = True, margins_name ='Total Shots/Vib')
                                Number_Shots_Per_vib_Pivot    = pd.DataFrame(Number_Shots_Per_vib_Pivot)
                                Number_Shots_Per_vib_Pivot    = Number_Shots_Per_vib_Pivot.reset_index(drop=False)
                                Number_Shots_Per_vib_Pivot.rename(columns={'Production_Day':"Production Day ↓ ", 'Total Shots/Vib':"TotalShots/Day"}, inplace = True)
                                Number_Shots_Per_vib_Pivot["Profile ID : → "]     = Number_Shots_Per_vib_Pivot.shape[0]*[" "]
                                Number_Shots_Per_vib_Pivot    = Number_Shots_Per_vib_Pivot.reset_index(drop=True)
                                cols = Number_Shots_Per_vib_Pivot.columns.tolist()
                                Cols_x = cols[-1:] + cols[1:-1]
                                Cols_y = cols[0:1]
                                Rearrange_Cols  = Cols_y + Cols_x
                                Number_Shots_Per_vib_Pivot = Number_Shots_Per_vib_Pivot[Rearrange_Cols]

                                ### Export Vib Production table report
                                def get_VibProductionSummary_Rep_datetime():
                                    return "Vib Production Summary -" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"
                                VibProductionSummary          = get_VibProductionSummary_Rep_datetime()        
                                outfile_VibProductionSummary  = "C:\\VibProductionReport\\" + VibProductionSummary        
                                XLSX_writer                   = pd.ExcelWriter(outfile_VibProductionSummary)
                                
                                Number_Shots_Per_vib_Pivot.to_excel(XLSX_writer,'VibProductionSummary',index=False, startrow=1)
                                Vib_Number.to_excel(XLSX_writer,'VibProductionSummary', index=False, startrow=0, startcol=1, header=False)
                                workbook       = XLSX_writer.book
                                worksheet      = XLSX_writer.sheets['VibProductionSummary']
                                header1 = '&L&G'+'&CEagle Canada Seismic Services ULC' + '\n' + '6806 Railway Street SE' + '\n' + 'Calgary, AB T2H 3A8' + '\n' +  'Ph: (403) 263-7770' +  '&R&U&18&"cambria, bold"Vib Production Sweep Report' + '\n' +  'Date : &D'
                                worksheet.set_header(header1,{'image_left':'eagle logo.jpg'})
                                footer1 = ('&LDate : &D')
                                worksheet.set_footer(footer1)
                                worksheet.set_landscape()
                                worksheet.set_margins(0.6, 0.6, 1.6, 1.1)
                                worksheet.print_across()
                                worksheet.fit_to_pages(1, 1)                                    
                                worksheet.set_paper(5)
                                worksheet.set_start_page(1)
                                worksheet.hide_gridlines(1)
                                worksheet.set_page_view()
                                workbook.formats[0].set_align('center')
                                workbook.formats[0].set_font_size(11)
                                workbook.formats[0].set_bold(True)
                                workbook.formats[0].set_border(1)
                                worksheet.set_column(0, 0, 20)   
                                worksheet.set_column(1, 1, 13)   
                                worksheet.set_column(LengthVibProfileIDList + 2, LengthVibProfileIDList + 3, 16)
                                worksheet.write('A1', Total_ProductionShots)
                                XLSX_writer.save()
                                XLSX_writer.close()
                                tkinter.messagebox.showinfo("Vib Production Report Export Message","Vib Production Summary Excel Report And Bar Plot Report Saved in : Location C:\\VibProductionReport")


                                #### Export Vib Production Plot Report
                                iPlot = tkinter.messagebox.askyesno("Vib Production Summary Excel Report And Bar Plot Report Save In Different Location",
                                                                    "Do You Like to Save Vib Production Summary Excel Report And Bar Plot Report Report in Different Location?")
                                if iPlot >0:
                                    def get_Plot_datetime():
                                        return "Vib Production Summary Plot Bar - " + datetime.now().strftime("%Y%m%d-%H%M%S") + ".pdf"

                                    def get_PlotBarh_datetime():
                                        return "Vib Production Summary Plot Barh - " + datetime.now().strftime("%Y%m%d-%H%M%S") + ".pdf"

                                    def get_VibProductionSummary_Rep_datetime():
                                        return "Vib Production Summary -" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"


                                    Plot_datetime_apply         = get_Plot_datetime()
                                    Barh_Plot_datetime_apply    = get_PlotBarh_datetime()
                                    VibProductionSummary        = get_VibProductionSummary_Rep_datetime()

                                    
                                    root                        = Tk()
                                    root.filename               =  tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Save Plot File As PDF ",
                                                                                                 filetypes = (("PDF files","*.pdf"),("all files","*.*")))
                                    if len(root.filename) >0:
                                        Plot_path                    = root.filename + Plot_datetime_apply
                                        BarH_Plot_path               = root.filename + Barh_Plot_datetime_apply
                                        outfile_VibProductionSummary = root.filename + VibProductionSummary
                                        pdfBarH = matplotlib.backends.backend_pdf.PdfPages(BarH_Plot_path)
                                        for fig in range(BarhInitial, plt.gcf().number + 1):
                                            pdfBarH.savefig( fig, dpi=10, orientation='portrait', bbox_inches='tight',metadata={'Title': 'Vib Production Summary'} )
                                        pdfBarH.close()
                                        pdf = matplotlib.backends.backend_pdf.PdfPages(Plot_path)
                                        for fig in range(1, ((plt.gcf().number + 1)-BarhInitial + 1)):
                                            pdf.savefig( fig, dpi=10, orientation='portrait', bbox_inches='tight',metadata={'Title': 'Vib Production Summary'} )
                                        pdf.close()                    
                                        plt.close('all')                                                                                             
                                        XLSX_writer = pd.ExcelWriter(outfile_VibProductionSummary)
                                        Number_Shots_Per_vib_Pivot.to_excel(XLSX_writer,'VibProductionSummary',index=False, startrow=1)
                                        Vib_Number.to_excel(XLSX_writer,'VibProductionSummary', index=False, startrow=0, startcol=1, header=False)
                                        workbook       = XLSX_writer.book
                                        worksheet      = XLSX_writer.sheets['VibProductionSummary']
                                        header1 = '&L&G'+'&CEagle Canada Seismic Services ULC' + '\n' + '6806 Railway Street SE' + '\n' + 'Calgary, AB T2H 3A8' + '\n' +  'Ph: (403) 263-7770' +  '&R&U&18&"cambria, bold"Vib Production Sweep Report' + '\n' +  'Date : &D'
                                        worksheet.set_header(header1,{'image_left':'eagle logo.jpg'})
                                        footer1 = ('&LDate : &D')
                                        worksheet.set_footer(footer1)
                                        worksheet.set_landscape()
                                        worksheet.set_margins(0.6, 0.6, 1.6, 1.1)
                                        worksheet.print_across()
                                        worksheet.fit_to_pages(1, 1)                                    
                                        worksheet.set_paper(5)
                                        worksheet.set_start_page(1)
                                        worksheet.hide_gridlines(1)
                                        worksheet.set_page_view()
                                        workbook.formats[0].set_align('center')
                                        workbook.formats[0].set_font_size(11)
                                        workbook.formats[0].set_bold(True)
                                        workbook.formats[0].set_border(1)
                                        worksheet.set_column(0, 0, 20)   
                                        worksheet.set_column(1, 1, 13)   
                                        worksheet.set_column(LengthVibProfileIDList + 2, LengthVibProfileIDList + 3, 16)
                                        worksheet.write('A1', Total_ProductionShots)
                                        XLSX_writer.save()
                                        XLSX_writer.close()
                                        tkinter.messagebox.showinfo("Vib Production Report Export Message"," Vib Production Bar Plot Report Saved As PDF And Summary Report Saved As Excel")
                                        root.destroy()

                                    else:
                                        tkinter.messagebox.showinfo("Vib Production Report Export Message","Please Select Vib Production Report File Name To Save")

                            else:
                                tkinter.messagebox.showinfo("Input File Error Message","Please Select File PSS Files to Process Report ")

                    ReadwindowVibInfo.Close()                                

                except ValueError:
                    tkinter.messagebox.showinfo("Error Message","Please Input Vib Profile ID And Number Correctly")

            except ValueError:
                tkinter.messagebox.showerror("TotalNumberVibes Error Message","TotalNumberVibes Must Be Number")
            
    WindowRead.Close()        

                


