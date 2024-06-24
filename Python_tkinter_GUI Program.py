import pyodbc 
import pandas as pd
import tkinter as tk
import re
from tkinter import filedialog, messagebox, ttk


# initalise the tkinter GUI
root = tk.Tk()

root.geometry("450x350") # set the root dimensions

root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0)

root.title("SMT - MPS")

# Frame for open file dialog
file_frame = tk.LabelFrame(root, background="white", text="SMT - Manpower(Monthly)/(Weekly) Uploader")
file_frame.place(height=350, width=450, rely=0.006, relx=0)


labelTop = tk.Label(root, background="white", text = "Step 1 - Select your SQL table:")
labelTop.grid(column=1, row=0, ipadx=25, pady=38, sticky="w")

comboExample = ttk.Combobox(file_frame, values=["MPS_Month_Raw", "MPS_Week_Raw"],width = 15)                        
comboExample.place(rely=0.055, relx=0.6)

# Buttons_Select your excel file
Label1 = tk.Label(root, background="white", text="Step 2 - Select your excel file:")
Label1.grid(column=1, row=1, ipadx=25, pady=38, sticky="w")

button1 = tk.Button(file_frame, text="Browse A File", width = 17, command=lambda: File_dialog())
button1.place(rely=0.339, relx=0.6)

# Buttons_Upload file to SQL Server
Label2 = tk.Label(root, background="white",text="Step 3 - Upload file to SQL Server:")
Label2.grid(column=1, row=3, ipadx=25, pady=38, sticky="w")
button2 = tk.Button(file_frame, text="Load File", width = 17, command=lambda: Load_excel_data())
button2.place(rely=0.638, relx=0.6)

# The file/file path text
label_file = ttk.Label(file_frame, background="white",text="No File Selected", font=("Time New Roman", 11))
label_file.place(rely=0.85, relx=0.04)

# def_File_dialog():
def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xls files", "*.xls"),("xlsx files", "*.xlsx"),("csv files", "*.csv"),("All Files", "*.*")))
    label_file["text"] = filename
    return None

# Read Excel Data
def Load_excel_data():
    
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df1 = pd.read_csv(excel_filename, keep_default_na = False)
            re1 = re.search('Manp(.*?)Based_on',excel_filename).group(1)
            re2 = re.search('on_MPS_(.*?).cs', excel_filename).group(1)
            re3 = re.search(r'\d+',re2).group(0)
        else:
            df1 = pd.read_excel(excel_filename, keep_default_na = False)
            re1 = re.search('Manp(.*?)Based_on',excel_filename).group(1)
            re2 = re.search('on_MPS_(.*?).xl', excel_filename).group(1)
            re3 = re.search(r'\d+',re2).group(0)

        conn = pyodbc.connect(
        Trusted_Connection='no',
        DRIVER='{ODBC Driver 17 for SQL Server}',
        server='Server name',
        DATABASE='Database name',
        UID='User ID',
        PWD='Password')

        cursor = conn.cursor()
    

#If file name is 'MPS_Month_Raw', insert column
        if str(comboExample.get()) == "MPS_Month_Raw":
            if re1 == "ower(Monthly)_":
                df = pd.DataFrame(df1)
                df.insert(0, 'Date', re3)
                df.columns.values[1] = "Org"
                df.columns.values[2] = "Item"
                df.columns.values[3] = "Item_Desc"
                df.columns.values[4] = "Cat1"
                df.columns.values[5] = "Cat2"
                df.columns.values[6] = "Cat3"
                df.columns.values[7] = "ReferenceRsc"
                df.columns.values[8] = "RscType"
                df.columns.values[9] = "OSPFlag"
                df.columns.values[10] = "MapFlag"
                df.columns.values[11] = "UPH"
                df.columns.values[12] = "Usage"
                df.columns.values[13] = "ReferenceItem"
                df.columns.values[14] = "ReferenceLevel"
                df.columns.values[15] = "ManHrsPcsERPActual"
                df.columns.values[16] = "ManHrsPcsERPSTD"
                df.columns.values[17] = "ManHrsPcsMAPSTD"
                df.columns.values[18] = "MainProcess"
                df.columns.values[19] = "Source"
                df.columns.values[20] = "Type"
                df.columns.values[21] = "Month_1"
                df.columns.values[22] = "Month_2"
                df.columns.values[23] = "Month_3"
                df.columns.values[24] = "Month_4"
                df.columns.values[25] = "Month_5"
                df.columns.values[26] = "Month_6"
                df.columns.values[27] = "Month_7"
                df.columns.values[28] = "Month_8"
                df.columns.values[29] = "Month_9"
                df.columns.values[30] = "Month_10"
                df.columns.values[31] = "Month_11"
                df.columns.values[32] = "Month_12"

# If not correct file, show the error message
            else:
                tk.messagebox.showerror("Information", "The file you have chosen is not Manpower(Monthly)")

# If file name correct, insert  data to sql
            for row in df.itertuples():
                cursor.execute('''
                                    INSERT INTO MPS_Month_Raw (Date, Org, Item, Item_Desc, Cat1, Cat2, Cat3, ReferenceRsc, RscType, OSPFlag, MapFlag, UPH, Usage, ReferenceItem, ReferenceLevel, ManHrsPcsERPActual, ManHrsPcsERPSTD,
                                    ManHrsPcsMAPSTD, MainProcess, Source, Type, Month_1, Month_2, Month_3, Month_4, Month_5, Month_6, Month_7, Month_8, Month_9, Month_10, Month_11, Month_12)
                                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                                    ''',
                                        row.Date,
                                        row.Org,
                                        row.Item,
                                        row.Item_Desc,
                                        row.Cat1,
                                        row.Cat2,
                                        row.Cat3,
                                        row.ReferenceRsc,
                                        row.RscType,
                                        row.OSPFlag,
                                        row.MapFlag,
                                        row.UPH,
                                        row.Usage,
                                        row.ReferenceItem,
                                        row.ReferenceLevel,
                                        row.ManHrsPcsERPActual,
                                        row.ManHrsPcsERPSTD,
                                        row.ManHrsPcsMAPSTD,
                                        row.MainProcess,
                                        row.Source,
                                        row.Type,
                                        row.Month_1,
                                        row.Month_2,
                                        row.Month_3,
                                        row.Month_4,
                                        row.Month_5,
                                        row.Month_6,
                                        row.Month_7,
                                        row.Month_8,
                                        row.Month_9,
                                        row.Month_10,
                                        row.Month_11,
                                        row.Month_12)

            
            conn.commit()
            cursor.close
            tk.messagebox.showinfo("Upload Successfully", "Upload Successfully")
            return None



#If file name is 'MPS_Week_Raw', insert column
        elif str(comboExample.get()) == "MPS_Week_Raw":
            if re1 == "ower(Weekly)_":
                df2 = pd.DataFrame(df1)
                df2.insert(0, 'Date', re3)
                df2.columns.values[1] = "Org"
                df2.columns.values[2] = "Item"
                df2.columns.values[3] = "Item_Desc"
                df2.columns.values[4] = "Cat1"
                df2.columns.values[5] = "Cat2"
                df2.columns.values[6] = "Cat3"
                df2.columns.values[7] = "ReferenceRsc"
                df2.columns.values[8] = "RscType"
                df2.columns.values[9] = "OSPFlag"
                df2.columns.values[10] = "MapFlag"
                df2.columns.values[11] = "UPH"
                df2.columns.values[12] = "Usage"
                df2.columns.values[13] = "ReferenceItem"
                df2.columns.values[14] = "ReferenceLevel"
                df2.columns.values[15] = "ManHrsPcsERPActual"
                df2.columns.values[16] = "ManHrsPcsERPSTD"
                df2.columns.values[17] = "ManHrsPcsMAPSTD"
                df2.columns.values[18] = "MainProcess"
                df2.columns.values[19] = "Source"
                df2.columns.values[20] = "Type"
                df2.columns.values[21] = "Week_1"
                df2.columns.values[22] = "Week_2"
                df2.columns.values[23] = "Week_3"
                df2.columns.values[24] = "Week_4"
                df2.columns.values[25] = "Week_5"
                df2.columns.values[26] = "Week_6"
                df2.columns.values[27] = "Week_7"
                df2.columns.values[28] = "Week_8"
                df2.columns.values[29] = "Week_9"
                df2.columns.values[30] = "Week_10"
                df2.columns.values[31] = "Week_11"
                df2.columns.values[32] = "Week_12"
                df2.columns.values[33] = "Week_13"
                df2.columns.values[34] = "Week_14"
                df2.columns.values[35] = "Week_15"
                df2.columns.values[36] = "Week_16"
                df2.columns.values[37] = "Week_17"
                df2.columns.values[38] = "Week_18"
                df2.columns.values[39] = "Week_19"
                df2.columns.values[40] = "Week_20"
                df2.columns.values[41] = "Week_21"
                df2.columns.values[42] = "Week_22"
                df2.columns.values[43] = "Week_23"
                df2.columns.values[44] = "Week_24"
                df2.columns.values[45] = "Week_25"
                df2.columns.values[46] = "Week_26"
                df2.columns.values[47] = "Week_27"
                df2.columns.values[48] = "Week_28"
                df2.columns.values[49] = "Week_29"
                df2.columns.values[50] = "Week_30"
                df2.columns.values[51] = "Week_31"
                df2.columns.values[52] = "Week_32"
                df2.columns.values[53] = "Week_33"
                df2.columns.values[54] = "Week_34"
                df2.columns.values[55] = "Week_35"
                df2.columns.values[56] = "Week_36"
                df2.columns.values[57] = "Week_37"
                df2.columns.values[58] = "Week_38"
                df2.columns.values[59] = "Week_39"
                df2.columns.values[60] = "Week_40"
                df2.columns.values[61] = "Week_41"
                df2.columns.values[62] = "Week_42"
                df2.columns.values[63] = "Week_43"
                df2.columns.values[64] = "Week_44"
                df2.columns.values[65] = "Week_45"
                df2.columns.values[66] = "Week_46"
                df2.columns.values[67] = "Week_47"
                df2.columns.values[68] = "Week_48"
                df2.columns.values[69] = "Week_49"
                df2.columns.values[70] = "Week_50"
                df2.columns.values[71] = "Week_51"
                df2.columns.values[72] = "Week_52"

            else:
                tk.messagebox.showerror("Information", "The file you have chosen is not Manpower(Weekly)")

# If file name correct, insert  data to sql
            for row in df2.itertuples():
                cursor.execute('''
                            INSERT INTO MPS_Week_Raw (Date, Org, Item, Item_Desc, Cat1, Cat2, Cat3, ReferenceRsc, RscType, OSPFlag, MapFlag, UPH, Usage, ReferenceItem, ReferenceLevel, ManHrsPcsERPActual, ManHrsPcsERPSTD,
                            ManHrsPcsMAPSTD, MainProcess, Source, Type, Week_1,	Week_2,	Week_3,	Week_4,	Week_5,	Week_6,	Week_7,	Week_8,	Week_9,	Week_10, Week_11, Week_12,
                            Week_13, Week_14, Week_15, Week_16,	Week_17, Week_18, Week_19, Week_20,	Week_21, Week_22, Week_23, Week_24, Week_25, Week_26, Week_27, Week_28,
                            Week_29, Week_30, Week_31, Week_32,	Week_33, Week_34, Week_35, Week_36, Week_37, Week_38, Week_39, Week_40,	Week_41, Week_42, Week_43, Week_44,	
                            Week_45, Week_46, Week_47, Week_48, Week_49, Week_50, Week_51, Week_52	)
                            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                            ''',
                                        row.Date,
                                        row.Org,
                                        row.Item,
                                        row.Item_Desc,
                                        row.Cat1,
                                        row.Cat2,
                                        row.Cat3,
                                        row.ReferenceRsc,
                                        row.RscType,
                                        row.OSPFlag,
                                        row.MapFlag,
                                        row.UPH,
                                        row.Usage,
                                        row.ReferenceItem,
                                        row.ReferenceLevel,
                                        row.ManHrsPcsERPActual,
                                        row.ManHrsPcsERPSTD,
                                        row.ManHrsPcsMAPSTD,
                                        row.MainProcess,
                                        row.Source,
                                        row.Type,
                                        row.Week_1,
                                        row.Week_2,
                                        row.Week_3,
                                        row.Week_4,
                                        row.Week_5,
                                        row.Week_6,
                                        row.Week_7,
                                        row.Week_8,
                                        row.Week_9,
                                        row.Week_10,
                                        row.Week_11,
                                        row.Week_12,
                                        row.Week_13,
                                        row.Week_14,
                                        row.Week_15,
                                        row.Week_16,
                                        row.Week_17,
                                        row.Week_18,
                                        row.Week_19,
                                        row.Week_20,
                                        row.Week_21,
                                        row.Week_22,
                                        row.Week_23,
                                        row.Week_24,
                                        row.Week_25,
                                        row.Week_26,
                                        row.Week_27,
                                        row.Week_28,
                                        row.Week_29,
                                        row.Week_30,
                                        row.Week_31,
                                        row.Week_32,
                                        row.Week_33,
                                        row.Week_34,
                                        row.Week_35,
                                        row.Week_36,
                                        row.Week_37,
                                        row.Week_38,
                                        row.Week_39,
                                        row.Week_40,
                                        row.Week_41,
                                        row.Week_42,
                                        row.Week_43,
                                        row.Week_44,
                                        row.Week_45,
                                        row.Week_46,
                                        row.Week_47,
                                        row.Week_48,
                                        row.Week_49,
                                        row.Week_50,
                                        row.Week_51,
                                        row.Week_52 )
            conn.commit()
            cursor.close
            tk.messagebox.showinfo("Upload Successfully", "Upload Successfully")
            return None

    # Error_message_popout
    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", "No file selected")
        return None

root.mainloop()