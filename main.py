import openpyxl as op
import os
import pandas as pd
import shutil
from tkinter import *
from tkinter import filedialog

vja_file_path=""
mumbai_file_path=""


def clear_file_paths():
    global vja_file_path
    global mumbai_file_path

    vja_path_label = Label(root,text = ' '*len(vja_file_path)*3)
    vja_path_label.config(font=("Arial", 12))
    vja_path_label.place(x=300,y=50)

    mumbai_path_label = Label(root,text = ' '*len(mumbai_file_path)*3)
    mumbai_path_label.config(font=("Arial", 12))
    mumbai_path_label.place(x=300,y=100)



def get_vja_file_path():
    global vja_file_path
    vja_file_path = filedialog.askopenfilename(initialdir = "/", title = "Select a File", filetypes = (("XLSX File", "*.xlsx*"),("CSV File", "*.csv*"),("Excel", "*.xls*"),("All files", "*.*"))) 

    vja_path_label = Label(root,text = vja_file_path)
    vja_path_label.config(font=("Arial", 12))
    vja_path_label.place(x=300,y=50)

    if os.path.isfile(vja_file_path):
        message_desc=Label(root,text="Vijayawada TCS File exists in the given path")
        message_desc.config(font=("Arial", 12),foreground="green")
        message_desc.place(x=20,y=300)
        message_desc.after(1000,message_desc.destroy)
    else:
        message_desc=Label(root,text="File does not exist in the given path")
        message_desc.config(font=("Arial", 12),foreground="red")
        message_desc.place(x=20,y=300)
        message_desc.after(1000,message_desc.destroy)

def get_mumbai_file_path():
    global mumbai_file_path
    
    mumbai_file_path = filedialog.askopenfilename(initialdir = "/", title = "Select a File", filetypes = (("XLSX File", "*.xlsx*"),("CSV File", "*.csv*"),("Excel", "*.xls*"),("All files", "*.*"))) 
    mumbai_path_label = Label(root,text = mumbai_file_path)
    mumbai_path_label.config(font=("Arial", 12))
    mumbai_path_label.place(x=300,y=100)

    if os.path.isfile(vja_file_path):
        message_desc=Label(root,text="Mumbai TCS File exists in the given path")
        message_desc.config(font=("Arial", 12),foreground="green")
        message_desc.place(x=20,y=300)
        message_desc.after(1000,message_desc.destroy)
    else:
        message_desc=Label(root,text="File does not exist in the given path")
        message_desc.config(font=("Arial", 12),foreground="red")
        message_desc.place(x=20,y=300)
        message_desc.after(1000,message_desc.destroy)


def generate_tcs_consolidation():

    global vja_file_path
    global mumbai_file_path
    master_pan_party_dict=dict()

    if os.path.exists(vja_file_path):
        vja_wb=op.load_workbook(vja_file_path)
        vja_df=pd.DataFrame(vja_wb.active.values)
        vja_df.columns=vja_df.iloc[4]
        vja_total_value=vja_df.iloc[-3]["Assessable Value"]
        vja_tcs_value=vja_df.iloc[-3]["TCS Rate"]
        vja_df=vja_df[5:-3]
    else:
        print("Accounts file missing")

    if os.path.exists(mumbai_file_path):
        mumbai_wb=op.load_workbook(mumbai_file_path)
        mumbai_df=pd.DataFrame(mumbai_wb.active.values)
        mumbai_df.columns=mumbai_df.iloc[3]
        mumbai_total_value=mumbai_df.iloc[-3]["Assessable Value"]
        mumbai_tcs_value=mumbai_df.iloc[-3]["TCS Rate"]
        mumbai_df=mumbai_df[4:-3]
    else:
        print("GSTR file missing")
    
    result_df = pd.concat([vja_df,mumbai_df])
    
    for index,row in result_df.iterrows():
        master_pan_party_dict[row["PAN"].strip()]=row["Party"]
    
    result_df=result_df.groupby(['PAN'])[["Assessable Value","TCS Rate"]].sum()
    result_df["PAN"]=result_df.index
    result_df["PAN"] = result_df["PAN"].apply(lambda x: x.strip())
    result_df["Assessable Value"]=result_df["Assessable Value"].apply(lambda x: round(x))
    result_df.reset_index(drop=True, inplace=True)
    result_df["Party Name"]=""
    result_df["Interest"]=0
    result_df["Total"]=0
    result_df["TCS Amount"]=0

    tcs_rate=tcs_rate_entry.get()
    if len(tcs_rate)==0:
        tcs_rate=0.01
    else:
        tcs_rate=float(tcs_rate)
    
    for index,row in result_df.iterrows():
        tcs_value=round((row["Assessable Value"]*tcs_rate)/100)
        result_df.loc[index,"Party Name"]=master_pan_party_dict[row["PAN"]]
        result_df.loc[index,"TCS Rate"]=tcs_rate
        result_df.loc[index,"Total"]=tcs_value
        result_df.loc[index,"TCS Amount"]=tcs_value

    
    result_df.rename(columns={'Assessable Value':'Amount','TCS Rate':'TCS %'}, inplace=True)
    result_df=result_df[["Party Name","PAN","Amount","TCS %","TCS Amount","Interest","Total"]]
    result_df=result_df.sort_values(by = 'Party Name') 
    total_dict={  
                "Party Name":"Total",
                "PAN":"",
                "TCS %":" ",
                "Amount":sum(result_df["Amount"]),
                "TCS Amount":sum(result_df["TCS Amount"]),
                "Interest":sum(result_df["Interest"]),
                "Total":sum(result_df["Total"])
            }
    
    row_dict={
                "Party Name":" ",
                "PAN":" ",
                "Amount":" ",
                "TCS %":" ",
                "TCS Amount":" ",
                "Interest":" ",
                "Total":" "
            }
    result_df=result_df.append(row_dict,ignore_index=True)
    result_df=result_df.append(total_dict,ignore_index=True)
    cwd=os.getcwd()
    output_file_path=os.path.join(cwd,'temp','tco_consolidated_output.csv')
    if os.path.exists(output_file_path):
        os.remove(output_file_path)

    result_df.to_csv(output_file_path,index=False)
    os.startfile(output_file_path)





if __name__=="__main__":

    root = Tk()
    root.title("Twills Clothing Pvt. Ltd.")
    root.geometry("550x350")

    header_label_one=Label(root,text="TCS Consolidation Tool",anchor="w")
    header_label_one.config(font=("Arial", 16))
    header_label_one.place(x=10,y=10)

    instruction_button = Button(root, text="Instructions")
    instruction_button.config(font=("Arial", 12))
    instruction_button.place(x=400,y=10)

    vja_label=Label(root,text="Vijayawada TCS File : ",font=("bold",10))
    vja_label.config(font=("Arial", 12))
    vja_label.place(x=10,y=50)

    vja_data_file = Button(root,text = "Choose File",command=get_vja_file_path)
    vja_data_file.config(font=("Arial", 12))
    vja_data_file.place(x=200,y=50)

    vja_path_label = Label(root,text = "")
    vja_path_label.config(font=("Arial", 12))
    vja_path_label.place(x=300,y=50)

    mumbai_label=Label(root,text="Mumbai TCS File : ",font=("bold",10))
    mumbai_label.config(font=("Arial", 12))
    mumbai_label.place(x=10,y=100)

    mumbai_data_file = Button(root,text = "Choose File",command=get_mumbai_file_path)
    mumbai_data_file.config(font=("Arial", 12))
    mumbai_data_file.place(x=200,y=100)

    mumbai_path_label = Label(root,text = "")
    mumbai_path_label.config(font=("Arial", 12))
    mumbai_path_label.place(x=300,y=100)

    tcs_rate_label=Label(root,text="TCS Rate (%) : ",font=("bold",10))
    tcs_rate_label.config(font=("Arial", 12))
    tcs_rate_label.place(x=10,y=150)

    tcs_rate_entry=Entry(root,width=8)
    tcs_rate_entry.config(font=("Arial", 12))
    tcs_rate_entry.place(x=200,y=150)

    button=Button(root,text="Consolidate Data",command=generate_tcs_consolidation)
    button.config(font=("Arial", 12))
    button.place(x=10,y=200)

    button=Button(root,text="Clear",command=clear_file_paths)
    button.config(font=("Arial", 12))
    button.place(x=180,y=200)

    message_label = Label(root,text = "Message :")
    message_label.config(font=("Arial", 12))
    message_label.place(x=10,y=250)

    message_label=Label(root,text="Welcome !!")
    message_label.config(font=("Arial", 12),fg="blue")
    message_label.place(x=20,y=300)
    message_label.after(1000,message_label.destroy)
    
    root.mainloop()

