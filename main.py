import calculation 
import dirutility as dutil
import departmentdata as college
import config
import reportutility as rutil

import traceback
import tkinter as tk
from datetime import datetime
from tkinter import ttk
from tkinter import filedialog, messagebox
from pprint import pprint

import os
import subprocess
import sqlite3
import pathlib
import shutil

import openpyxl
from dateutil import parser

# Constants
MIXED       = "MIXED"
HGT         = 2
FONT        = ('', 16)
order       = 1
inputs : list = []
PX, PY = (10, 10)
FIELD_SIZE = 50
DATERANGE = tuple(range(1, 32))
MONTHRANGE = ('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
              'Sep', 'Oct', 'Nov', 'Dec')
YEARRANGE = tuple(range(datetime.now().year, 1999, -1))
COMBINED    = "COMBINED"
debug_setting = config.DEBUG_MODE


def Browse_Files() -> str:

    input_filetypes = [
        ('Excel files', '*.xlsx *.xlsm *.xls'),
        ('CSV files', '*.csv'),
        ('All files', '*.*'),
    ]

    filename = filedialog.askopenfilename(parent=frame_survey_details,
                                        initialdir=os.getcwd(),
                                        title="Please select a file:",
                                        filetypes=input_filetypes)

    input_filepath.configure(state=tk.NORMAL)
    input_filepath.delete(0, tk.END)
    input_filepath.insert(0, filename)
    input_filepath.configure(state='readonly')
    return filename

def Upload_Action():
    global inputs
    
    # Picking date(day/month/year)& time in date variable
    date = f"{pick_day.get()} {pick_month.get()} {pick_year.get()} {timevalue.get()}"

    if debug_setting:
        given_timestamp = parser.parse(date or "01/01/2023")
        rawfilename  = input_filepath.get() or 'cas-sample-score-sheet.xlsx'
    else:
        rawfilename  = input_filepath.get()
        if rawfilename.strip() == '':
            messagebox.showerror("Error", "Choose a file before uploading!")
            return
        try:         
            given_timestamp = parser.parse(date)
        except Exception:
            messagebox.showerror("Error", 
                "Given datetime format is not correct!\n" +\
                "Example: 01/03/22 3:30pm")
            return

    filename = pathlib.Path(rawfilename).stem
    inputs = [
        input_surveyname.get().strip() or filename,
        given_timestamp,
        pick_institute.get().strip(), 
        pick_department.get().strip(),
        # This second rawfilename will be modified in Upload_Report() 
        # function, so we need two of these
        rawfilename,
        rawfilename, 
        datetime.now()
    ]
    
    # print(inputs)

def openinfowindow(parentwindow):
    FALLBACK_INFO = """
Developed by:

* Ashutosh Dubey\t\t(IMCA 2021 batch)

\t\t\t\t
Student at Acropolis FCA department under guidance of \t\t
Prof. Nitin Kulkarni & Kushagra Mehrotra\t(IMCA 2019 batch) .

For help and assistance email at: 
ashutoshdubey.ca21@acropolis.in
"""
    infowindow = tk.Toplevel(parentwindow)
    info = FALLBACK_INFO
    label_contributors = tk.Label(infowindow, text=info, justify='left')
    label_contributors.pack(ipadx=PX)

def createFooter(parentwindow):
    INFO_DEPT       = """Made by students at FCA department"""
    label_madeby    = tk.Label(parentwindow, text=INFO_DEPT)
    button_info     = tk.Button(parentwindow, text="?", 
                        command=lambda: openinfowindow(parentwindow))

    button_info.pack(side='right', ipadx=PX,pady=PY)
    label_madeby.pack(side='right', padx=PX,pady=PY)

    return button_info, label_madeby

def updateView(event=None):
    global order
    filtertxt = input_filter.get()
    filtertxt = f"%{filtertxt}%"
    # Since both sortorder and order variables do not get data from text-field
    # so string interpolation with them is secure.
    sortorder = str(comboBoxMap[pick_order.get()])
    # order     = 'ASC' if 
    # To remove all items from table
    for item in tree.get_children():
       tree.delete(item)

    # con1 = sqlite3.connect(config.DB)
    print(con)
    con1 = con or sqlite3.connect(config.DB)
    cur1 = con1.cursor()

    cur1.execute("""SELECT * FROM tblSurveySheets WHERE survey_name LIKE ? 
        or institute LIKE ? or department LIKE ? COLLATE NOCASE ORDER BY """ +
        F"{sortorder} {'ASC' if order==1 else 'DESC'}", (filtertxt,)*3)
    
    rows = cur1.fetchall()

    # For displaying All matching items
    for row in rows:
        
        displayname = f"({row[3] if row[3] not in ('', '...') else 'Unspecified'}" + \
                    f"{' '+str(row[4]) if row[4]!='' else ''}) {row[1]}"
        giventime   = parser.parse(row[2])
        uploadtime  = parser.parse(row[-1])
        ogpath      = row[-3]
        filename    = row[-2]
        primaryid   = int(row[0])
        giventime   = giventime.strftime("%d-%m-%Y %H:%M")
        uploadtime  = uploadtime.strftime("%d-%m-%Y %H:%M")
        tree.insert("", tk.END, values=(displayname, giventime, uploadtime,
                                        ogpath, filename, primaryid))

    if debug_setting:
        print("updating... ")
        
        print("..updated!")

def clearInputs():
    input_filepath.configure(state='normal')
    input_surveyname.delete(0, tk.END)
    input_filepath.delete(0, tk.END)
    input_time.delete(0, tk.END)
    input_filepath.configure(state='readonly')

def Upload_Report(data: list):

    dutil.create_raw_folder()
    cryptic_filename = dutil.upload_to_raw_folder(data[-2])
    # Renaming filename(second) with cryptic name
    data[-2] = cryptic_filename

    con1 = con or sqlite3.connect(config.DB)
    cur = con1.cursor()
    cur.execute("""INSERT INTO 
        tblSurveySheets(survey_name, survey_time, institute, department, 
        ogfile_path, file, upload_time) VALUES(?,?,?,?,?,?,?)""", data)
    rows = cur.fetchall()    
    print("rows: ", rows)
    con1.commit()

def Upload_Action():
    global inputs
    
    # Getting date(day/month/year)& time in date variable
    date = f"{pick_day.get()} {pick_month.get()} {pick_year.get()} {timevalue.get()}"

    if debug_setting:
        # For mendatory inputs
        given_timestamp = parser.parse(date or "01/01/2023")
        rawfilename  = input_filepath.get() or 'Stress Scale Sample Data.xlsx'
    else:
        rawfilename  = input_filepath.get()
        if rawfilename.strip() == '':
            messagebox.showerror("Error", "Choose a file before uploading!")
            return
        try:           
            print(date)
            given_timestamp = parser.parse(date)
        except Exception:
            messagebox.showerror("Error", 
                "Given datetime format is not correct!\n" +\
                "Example: 01/03/22 3:30pm")
            return

    filename = pathlib.Path(rawfilename).stem
    inputs = [
        input_surveyname.get().strip() or filename,
        given_timestamp,
        pick_institute.get().strip(), 
        pick_department.get().strip(),
        # This second rawfilename will be modified in Upload_Report() 
        # function, so we need two of these
        rawfilename,
        rawfilename, 
        datetime.now()
    ]
    
    Upload_Report(inputs)
    updateView()
    clearInputs()
    print(inputs)
def Change_Sort_Order(event=None):
    global order
    order = not order
    button_order.configure(text = f"{'^' if order else 'v'}")
    updateView()
    disableButtons()

def disableButtons(event=None):
    global curItemId
    button_generate.configure(state=tk.DISABLED, text="PROCESS")
    button_savecopy.configure(state=tk.DISABLED, text="DOWNLOAD")
    button_delete.configure(state=tk.DISABLED, text="DELETE SURVEY")
    button_viewsum.configure(state=tk.DISABLED)
    ogfilename.set('')

    for item in tree.selection():
        tree.selection_remove(item)
    # curItemId = tree.focus()
    curItemId = ''
    print("disableButtons event:", tree.item(curItemId))

def enableButtons(a=None):
    curItemId = ''
    print("type is ", type(tree.selection()), 
        "items are", tree.selection())
    
    tot_selected_items = len(tree.selection())

    if tot_selected_items > 0:
        curItemId = tree.selection()[0]

    print("event is", a, "and Item ID is", curItemId, type(curItemId), len(curItemId))
    
    button_generate['state'] = tk.NORMAL
    item : list = tree.item(curItemId)['values']
    
    print(item, "its length is", len(item))
    

    if tot_selected_items > 0:
        button_delete['state'] = tk.NORMAL
        
        if tot_selected_items == 1:
            filename         = pathlib.Path(item[-2]).stem
            uploadedfilename = pathlib.Path(item[-3]).name
            ogfilename.set(f'({uploadedfilename})')
            
            if (pathlib.Path(f'{config.REPORTSFOL}')/filename).is_dir():
                button_generate['text'] = "REPROCESS"
                button_delete['text']  = "DELETE SURVEY"
                button_savecopy['state'] = tk.NORMAL
                button_viewsum['state'] = tk.NORMAL
                
            else:
                button_generate['text'] = "PROCESS"
                button_savecopy['state'] = tk.DISABLED
                button_viewsum['state'] = tk.DISABLED
            
        else:
            uploadedfiles = []
            for itemId in tree.selection():
                filepath            = tree.item(itemId)['values'][-3]
                uploadedfilename    = pathlib.Path(filepath).name
                uploadedfiles.append(uploadedfilename)

            filelist = ", ".join(uploadedfiles)
            ogfilename.set(f'({filelist[:80]}' + ( "...)" 
                if len(filelist) >= 80 else ")" ))
            button_generate['text'] = "PROCESS ALL"
            button_delete['text']  = "DELETE ALL"
            button_savecopy['state'] = tk.DISABLED
            button_viewsum['state'] = tk.DISABLED
    else:
        disableButtons()

def enterTime(event=None):
    return datetime.now().strftime("%H:%M:%S")
    pass

def openSubprocess(e=None):
    print(itemId := tree.selection()[0])
    data : list = tree.item(itemId)['values']
    file : str  = data[-2]
    sts = subprocess.Popen(f"\"{config.RAWFOL}\\{file}\"", shell=True)
    return sts

def updateDeptBox(e=None):
    insti = pick_institute.get()
    pick_department['values']=college.INST_DEPT_MAP.get(insti, (COMBINED, ))
    pick_department.current(0)

def Clear_Filter(event=None):
    input_filter.delete(0, tk.END)
    updateView()

def log_error(msg: str, *msgs: str):
    dutil.create_safe_dir(config.LOGDIR)
    LOGEXT  = r".log.txt"
    LOGEXT  = r".log"
    filename = pathlib.Path(datetime.now().strftime("%Y%m%d") + LOGEXT)
    # Check if file with today's date exists
    logfile = config.LOGDIR / filename
    # If it does not, create it.
    if not logfile.exists():
        logfile.touch()
    # open file in append mode
    with logfile.open(mode='a', encoding='utf-8', errors='xmlcharrefreplace') as handler:
        # print msg to the file-handler
        # print(datetime.now().strftime('%d/%m/%Y %I:%M:%S %p - '), 
        #     file=handler, end='')
        handler.write(datetime.now().strftime('%d/%m/%Y %I:%M:%S %p'))
        handler.write(' - ERROR - ')
        # handler.write(msg + '\n')
        print(msg, file=handler, end='\n')
        for m in msgs:
            # handler.write(m + '\n')
            print(m, file=handler, end='\n')

def Perform_File_Operations(data_sheet: str, survey_id: int, survey_name: str):
    
    rootwindow.update_idletasks()
    
    fname = os.path.join(config.RAWFOL, data_sheet)
    if (data_sheet.endswith(('.xlsx', '.xlsm', '.xls'))):
        data = calculation.Parse_Excel_To_List(fname)
    elif (data_sheet.endswith('.csv')):
        data = calculation.Parse_Csv_To_List(fname)
    else:
        # If given unrecognized file format
        return
    Summary_sheet = calculation.process(data)
    if (debug_setting):
        oldtscdata = tscdata[:]
        del tscdata
        tscdata = oldtscdata[:1]
        tscdata.append([71] * len(tscdata[0]))
    # pprint(studata) 
        
    reportsdir = dutil.create_report_folder(data_sheet)
    
    rutil.Create_Summary(reportsdir ,Summary_sheet,survey_name)
 
def center_window(window, width=300, height=200):
    # get screen width and height
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # calculate position x and y coordinates
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    window.geometry('%dx%d+%d+%d' % (width, height, x, y))

def Create_Busy_Frame(parentwindow):
   
    bsframe = tk.Toplevel(parentwindow)
    
    center_window(bsframe, 300, 100)
    bsframe.title("Wait...")
    label = tk.Label(bsframe, text="Generating reports...", anchor='center')
    
    label.pack(anchor='center', fill='both', expand=1)
    return bsframe
    pass

def Generate_Action():
    busyframe = Create_Busy_Frame(rootwindow)
    busyframe.update()
    busyframe.grab_set()
    rootwindow.withdraw()
    
    for treeItemId in tree.selection():
        
        print( treedata := tree.item(treeItemId)['values'] )
        item_id = int(treedata[-1])
        survey_name = treedata[0]
        
        con1 = con or sqlite3.connect(config.DB)
        cur = con1.cursor()
        fname = cur.execute("""SELECT file FROM tblSurveySheets WHERE id=?""", 
                            (str(item_id),) ).fetchone()[0]
                            
        print(fname)
        print()
        try:
            Perform_File_Operations(fname, item_id, survey_name)
        except Exception as e:
            busyframe.grab_release()
            rootwindow.deiconify()
            busyframe.destroy()
            # logging.error(traceback.format_exc())
            # log_error(traceback.format_exc())
            # log_error(str(e))
            log_error(e, traceback.format_exc())
            print("ERRORRRR: ", e)
            messagebox.showerror("Error", "Unable to generate reports.")
            raise e
    busyframe.grab_release()
    rootwindow.deiconify()
    busyframe.destroy()
    messagebox.showinfo("Done", "Reports generated!")
    disableButtons()

def Open_Associated_Summary():
    treeItemId  = tree.focus()
    treestorage = tree.item(treeItemId)['values']
    survey_name = treestorage[0]
    outputfol   = pathlib.Path(treestorage[-2]).stem
    reportspath = pathlib.Path(config.REPORTSFOL) / outputfol
    if (reportspath.exists()):
        summary_sheet_path = reportspath / rutil.getsummaryname(survey_name)
        print(f'{summary_sheet_path=}')
        print(str(summary_sheet_path)) 
        # subprocess.run(['start', 'excel.exe',f'"{str(summary_sheet_path)}"'], shell=True)
        subprocess.Popen(f'start excel.exe "{str(summary_sheet_path)}"', shell=True)

def Copy_Summary():
    dest    = filedialog.askdirectory(parent=rootwindow,
                                initialdir='/',
                                title="Please select a folder:")
    treedata        = tree.item(tree.focus())['values']
    selecteditemid  = tree.focus()
    outputfol       = pathlib.Path(treedata[-2]).stem
    survey_name     = treedata[0]
    reportspath     = pathlib.Path(config.REPORTSFOL) / outputfol
    if pathlib.Path(directory:=os.path.join(config.REPORTSFOL, outputfol)).is_dir():
        file        = rutil.getsummaryname(survey_name)
        srcpath     = pathlib.Path(directory) / file
        destpath    = pathlib.Path(dest) / file
       
        if destpath.exists():
            
            count = 1
            while (pathlib.Path(dest) / (file.stem + f" ({count})" + file.suffix)).exists():
                count += 1
            shutil.copy2(srcpath, (pathlib.Path(dest) / (file.stem + f" ({count})" + file.suffix)))
        else:
            shutil.copy2(srcpath, dest)
            # shutil.copy2(srcfile, dest)
        
        dest = dest.translate({ord('/'):'\\'})
        print("Summary sheet copies to desination!")
        print(dest)
        # messagebox.showinfo("Download complete!", "Download complete!!")
        answer = messagebox.askyesno("Download complete", 
                    "Would you like to open the folder?")
        if (answer == True):
            subprocess.run(["explorer", dest])

def Delete_Reports():
    """ for all items in tree.selection():
            fetch file name from either treeview or config.DB
            if folder(reports/file-name).exists
                empty it then delete it
            try:
                delete folder(raw)/filename
            finally:
                remove item from treeview as well
            delete row from config.DB table tblSurveySheets

            delete from config.DB as well
    """
    selected_items = tree.selection()

    # Message to confirm
    answer = messagebox.askyesno("Confirm deletion?", "Do you want to delete this survey and associated Sheet?")
    print(answer)

    if answer == True:
        print(selected_items)
        
        for itemid in selected_items:
            surveyid = tree.item(itemid)['values'][-1]
            con1 = con or sqlite3.connect(config.DB)
            cur = con1.cursor()
            cur.execute("""SELECT file from tblSurveySheets WHERE id=?""", 
                (surveyid,))
            filename = pathlib.Path(cur.fetchone()[0])
           
            if (targetdir:=pathlib.Path(config.REPORTSFOL)/filename.stem).is_dir():
                shutil.rmtree(targetdir)
                pass
            
            if (targetfile:=pathlib.Path(config.RAWFOL)/filename).exists():
                targetfile.unlink()
                pass

            cur.execute("""DELETE FROM tblSurveySheets WHERE id=?""",
                (surveyid,))
            
            con1.commit()
            tree.delete(itemid)
            print("Deletion done!")

if __name__ == '__main__':

    # For making Folder and creating Tables (Database)
    if not os.path.exists(config.DBLOC):
        os.mkdir(config.DBLOC)
    con = sqlite3.connect(config.DB)
    con.execute("PRAGMA foreign_keys = ON")
    cur = con.cursor()

    cur.execute("""CREATE TABLE IF NOT EXISTS tblSurveySheets(
        id INTEGER PRIMARY KEY,
        survey_name VARCHAR,
        survey_time TIMESTAMP,
        institute VARCHAR,
        department VARCHAR,
        ogfile_path VARCHAR,
        file VARCHAR,
        upload_time TIMESTAMP)""")


    # Creating the application window
    rootwindow = tk.Tk()
    rootwindow.title("Stress Meter")
    rootwindow.resizable(False, False)
    center_window(rootwindow,1160,560)

    frame_datetime = tk.Frame(rootwindow, width=FIELD_SIZE)

    institute_values = college.institutes
    department_value = college.INST_DEPT_MAP.get(institute_values[0], (COMBINED,))

    sort_keys = ('Upload Time (Default)', 'Survey Name', 'Survey Time',
                        'Institute', 'Department')
    db_fields   = ('upload_time', 'survey_name', 'survey_time', 
                        'institute', 'department')
    
    comboBoxMap = dict(zip(sort_keys, db_fields))
    timevalue   = tk.StringVar()
    timevalue.set('')

    # Creating Frames
    frame_survey_details = tk.LabelFrame(rootwindow,text="Survey Details ",width=FIELD_SIZE)
    frame_survey_details.grid(column=0,row=0,padx=PX,pady=(PY,0),sticky=tk.W+tk.E)


    frame_datetime = tk.Frame(frame_survey_details, width=FIELD_SIZE)
    frame_datetime.grid(column=0, row=5, rows=2, pady=(PY*2, 0), 
            sticky=tk.W+tk.E)
    
    frame_Action = tk.LabelFrame(rootwindow,text="Operations ",width=FIELD_SIZE)
    frame_Action.grid(column=6,row=0,padx=PX,pady=(PY,0),sticky=tk.N+tk.W+tk.E+tk.S)

    frame_footer = tk.Frame(rootwindow)
    frame_footer.grid(row=999, column=0, columns=999, sticky=tk.E+tk.W)
    createFooter(frame_footer)

    # Creating the lables, textboxs and buttons for Frame - Survey Details 

    label_filepath = tk.Label(frame_survey_details, text="Browse input file *", anchor="w")
    input_filepath = tk.Entry(frame_survey_details, width=FIELD_SIZE)
    button_browse = tk.Button(frame_survey_details,
                              text="Browse",command=Browse_Files,height= HGT)

    label_surveyname = tk.Label(frame_survey_details, text="Survey name", anchor="w")
    input_surveyname = tk.Entry(frame_survey_details, width=FIELD_SIZE)

    label_date = ttk.LabelFrame(frame_datetime,
                                text="Pick date of survey:*", width=FIELD_SIZE//2)
    pick_day = ttk.Combobox(label_date, values=DATERANGE,
                            state='readonly', width='3')
    pick_month = ttk.Combobox(label_date, values=MONTHRANGE,
                              state='readonly', width='4')
    pick_year = ttk.Combobox(label_date, values=YEARRANGE,
                             state='readonly', width='5')

    label_time = ttk.LabelFrame(frame_datetime,
                                text="Time of survey (optional):")
    timevalue = tk.StringVar()
    timevalue.set('')
    input_time = tk.Entry(label_time, textvariable=timevalue)

    label_institute = tk.Label(frame_survey_details, text="Institute", anchor='w')
    institute_values = college.institutes
    pick_institute = ttk.Combobox(frame_survey_details, width=FIELD_SIZE-3,
                                  text="Institute", values=institute_values, state='readonly')

    label_department = tk.Label(frame_survey_details, text="Department", anchor='w')
    department_value = college.INST_DEPT_MAP.get(institute_values[0], (COMBINED,))
    pick_department = ttk.Combobox(frame_survey_details, width=FIELD_SIZE-3,
                                   text="Department", values=department_value, state='readonly')

    button_upload = tk.Button(frame_survey_details, 
                            text = 'Upload form',command=Upload_Action,height= HGT)


    # Positioning the lables, textboxs and buttons for Frame - Survey Details 

    label_filepath.grid(column=0, row=0, pady=(PY),  sticky=tk.W+tk.E)
    input_filepath.grid(column=0, row=1, sticky=tk.W+tk.E)
    button_browse.grid(column=0, row=2, pady=(PY), padx=PX, sticky=tk.W+tk.E)

    label_surveyname.grid(column=0, row=3, pady=(PY, 0), sticky=tk.W+tk.E)
    input_surveyname.grid(column=0, row=4, sticky=tk.W+tk.E)

    label_date.pack(side='left', expand=1, fill='y')
    pick_day.pack(side='left', padx=(PX//2), pady=PY)
    pick_month.pack(side='left', padx=(PX//2), pady=PY)
    pick_year.pack(side='left', padx=(PX//2), pady=PY)
    pick_day.current(0)
    pick_month.current(0)
    pick_year.current(0)

    label_time.pack(side='right', expand=1, fill='y')
    input_time.pack(side='left', padx=PX, pady=PY)

    label_institute.grid(column=0, row=7, pady=(PY,0), sticky=tk.W+tk.E)
    pick_institute.grid(column=0, row=8, sticky=tk.W+tk.E)
    pick_institute.current(0)

    label_department.grid(column=0, row=9, sticky=tk.W+tk.E)
    pick_department.grid(column=0, row=10, sticky=tk.W+tk.E,pady=PY)
    pick_department.current(0)

    button_upload.grid(column=0, row=13, pady=PX, padx=PX, sticky=tk.W+tk.E)

    # For getting path of input file in input_filepath(read only)
    input_filepath.insert(0, "")
    input_filepath.configure(state='readonly')

    # For Showing Uploaded Files 
    tree = ttk.Treeview(rootwindow, columns=("surveyname","date",
                "uploadtime","ogfilepath","filename","pk"), show='headings', 
                displaycolumns=(0,1,2))

    tree.heading("surveyname",  text="Survey name", anchor="w")
    tree.heading("date",        text="Survey Timestamp", anchor="center")
    tree.heading("uploadtime",  text="Upload Timestamp", anchor="center")
    tree.heading("pk",          text="ID", anchor="w")

    tree.column("surveyname", stretch=True, minwidth=250)
    tree.column("date",         stretch=True, minwidth=100, width=150, 
                    anchor='center')
    tree.column("uploadtime",   stretch=True, minwidth=100, width=200, 
                    anchor='center')

    tree.rowconfigure(0, weight=1)
    tree.grid(column=2, row=0, rowspan=14, columns=4, 
                    sticky=tk.N+tk.W+tk.E+tk.S, padx=PX, pady=PY)

    ttk.Separator(rootwindow,
                orient=tk.HORIZONTAL, 
                style='TSeparator',
                cursor='man',
                ).grid(column=0, row=998, columns=999, sticky=tk.E+tk.W)


    # Creating the lables, textboxs and buttons for Frame - Operation
    button_clearbox = tk.Button(frame_Action,
                        text = 'clear filter', command=Clear_Filter)

    label_filter    = tk.Label(frame_Action, text="Filter:", anchor='w')
    input_filter    = tk.Entry(frame_Action, width=17)

    button_generate = tk.Button(frame_Action, 
                        text = 'Process',  state='disabled',height= HGT,
                        command=Generate_Action)

    button_viewsum  = tk.Button(frame_Action,
                        text = 'VIEW SUMMARY SHEET',      state='disabled',height= HGT,
                        command=Open_Associated_Summary)

    button_savecopy  = tk.Button(frame_Action, 
                        text = 'DOWNLOAD SUMMARY SHEET',  state='disabled',height= HGT,
                        command=Copy_Summary)

    button_delete   = tk.Button(frame_Action, 
                        text = 'DELETE SURVEY',     state='disabled',height= HGT,
                        command=Delete_Reports)

    label_sortby    = tk.Label(frame_Action, text="Sort by:", anchor='w')

    button_order    = tk.Button(frame_Action,
                        text = f"{'^' if order else 'v'}", 
                        command= Change_Sort_Order)
    
    sort_keys = ('Upload Time (Default)', 'Survey Name', 'Survey Time',
                        'Institute', 'Department')
    pick_order      = ttk.Combobox(frame_Action, text="sort_by", 
        values=sort_keys, state='readonly')
    
    ogfilename = tk.StringVar()
    ogfilename.set('')

    label_associatedfile = tk.Label(rootwindow, textvariable=ogfilename, 
                                    anchor='w')
    
    # Positioning the lables, textboxs and buttons for Frame - Survey Details 
    button_clearbox.grid(column=7, row=0, sticky=tk.E, padx=PX, pady=0)

    input_filter.grid(column=7, row=2, sticky=tk.E, padx=PX)
    label_filter.grid(column=7, row=2, sticky=tk.W, padx=PX,pady=PY)

    button_generate.grid(column=7, row=3, sticky=tk.E+tk.W, padx=PX,pady=PY)
    
    button_viewsum.grid(column=7, row=4, sticky=tk.E+tk.W, padx=PX,pady=PY)

    button_savecopy.grid(column=7, row=5, sticky=tk.E+tk.W, padx=PX,pady=PY)
    
    button_delete.grid(column=7, row=6, sticky=tk.E+tk.W, padx=PX,pady=PY)

    label_sortby.grid(column=7, row=12, sticky=tk.W, padx=PX)
    button_order.grid(column=7, row=12, sticky=tk.E, padx=PX, ipadx=PX/2)
  
    pick_order.grid(column=7, row=13, sticky=tk.W, padx=PX,pady=PY)
    pick_order.current(0)
    

    label_associatedfile.grid(column=2, row=15, columns=2, 
                            sticky=tk.N+tk.W+tk.E+tk.S, padx=PX, pady=PY)
    

    tree.bind('<ButtonRelease-1>', enableButtons)
    
    tree.bind('<<TreeviewSelect>>', enableButtons)
    tree.bind('<FocusOut>', disableButtons)
    
    tree.bind('<<TreeviewOpen>>', openSubprocess)
    input_filter.bind('<Return>', updateView)
    
    input_filter.bind('<KeyRelease>', updateView)
    
    input_time.bind('<Triple-Button-1>', 
                    lambda e:timevalue.set(enterTime()))
   
    pick_order.bind("<<ComboboxSelected>>", updateView)
    pick_institute.bind("<<ComboboxSelected>>", updateDeptBox)

    updateDeptBox()
    updateView()
    
    # Start gui
    rootwindow.mainloop()

    # Closing the connection
    con.close()
    print("Connection closed!")