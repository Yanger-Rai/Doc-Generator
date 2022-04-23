'''
version : Python 3.8.5
docx2pdf: 0.1.8
docxtpl : 0.15.2
pandas  : 1.2.4
owner: aanirudi@, vbikashr@
'''

#========imports=========
import os, sys
from docxtpl import DocxTemplate
from matplotlib import image
import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import messagebox
import warnings


#initiating UI
root = Tk()
root.title("France HVH TOE Tool")

#Try & except will ignore the "UerWarning" due to openpyxl lib 
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

#Style for the progress bar
s = ttk.Style()
s.theme_use('clam')
s.configure("white.Horizontal.TProgressbar", troughcolor = 'white', bordercolor = 'white', background = 'black')

#Place the UI at the center of the screen
window_height = 500
window_width = 900
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_cordinate = int((screen_width/2) - (window_width/2))
y_cordinate = int((screen_height/2) - (window_height/2))
root.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))

#=====glable variables=====
excel_name = StringVar()
template_name = StringVar()
output_folder_name = StringVar()

#needed for pyinstaller to package the code into exe
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

path = resource_path("page.png")
   
#This is the main page, well there is only one page :)
class main_page:
    def __init__(self, top=None):
        top.resizable(0, 0)
        top.title("France HVH TOE Tool")

        self.label = Label(top)
        self.label.place(relx=0, rely=0, width=900, height=500)
        self.img = PhotoImage(file=path)
        self.label.configure(image=self.img)
        
        #Excel Button 
        self.button1 = Button(top)
        self.button1.place(relx=0.09, rely=0.275, width=255, height=30)
        self.button1.configure(relief="flat")
        self.button1.configure(overrelief="flat")
        self.button1.configure(activebackground="#2C2B2B")
        self.button1.configure(cursor="hand2")
        self.button1.configure(foreground="#FFFFFF")
        self.button1.configure(background="#2C2B2B")
        self.button1.configure(font="-family {Poppins Regular} -size 10")
        self.button1.configure(borderwidth="0")
        self.button1.configure(text="""EXCEL""")
        self.button1.configure(command=lambda:self.select_file(excel_name))

        self.label1 = Label(top, textvariable=excel_name)
        self.label1.place(relx=0.4, rely=0.264, width=430, height=30)
        self.label1.configure(background="#FFFFFF")
        self.label1.configure(anchor="w")

        #Template Button 
        self.button2 = Button(top)
        self.button2.place(relx=0.09, rely=0.364, width=255, height=30)
        self.button2.configure(relief="flat")
        self.button2.configure(overrelief="flat")
        self.button2.configure(activebackground="#2C2B2B")
        self.button2.configure(cursor="hand2")
        self.button2.configure(foreground="#FFFFFF")
        self.button2.configure(background="#2C2B2B")
        self.button2.configure(font="-family {Poppins Regular} -size 10")
        self.button2.configure(borderwidth="0")
        self.button2.configure(text="""TEMPLATE""")
        self.button2.configure(command=lambda:self.select_file(template_name))

        self.label2 = Label(top, textvariable=template_name)
        self.label2.place(relx=0.4, rely=0.354, width=430, height=30)
        self.label2.configure(background="#FFFFFF")
        self.label2.configure(anchor="w")

        #Output Button 
        self.button3 = Button(top)
        self.button3.place(relx=0.09, rely=0.454, width=255, height=30)
        self.button3.configure(relief="flat")
        self.button3.configure(overrelief="flat")
        self.button3.configure(activebackground="#2C2B2B")
        self.button3.configure(cursor="hand2")
        self.button3.configure(foreground="#FFFFFF")
        self.button3.configure(background="#2C2B2B")
        self.button3.configure(font="-family {Poppins Regular} -size 10")
        self.button3.configure(borderwidth="0")
        self.button3.configure(text="""OUTPUT FOLDER""")
        self.button3.configure(command=lambda:self.select_dic(output_folder_name))

        self.label3 = Label(top, textvariable=output_folder_name)
        self.label3.place(relx=0.4, rely=0.444, width=430, height=30)
        self.label3.configure(background="#FFFFFF")
        self.label3.configure(anchor="w")
        
        #Generate Button 
        self.button4 = Button(top)
        self.button4.place(relx=0.09, rely=0.664, width=255, height=30)
        self.button4.configure(relief="flat")
        self.button4.configure(overrelief="flat")
        self.button4.configure(activebackground="#2C2B2B")
        self.button4.configure(cursor="hand2")
        self.button4.configure(foreground="#FFFFFF")
        self.button4.configure(background="#2C2B2B")
        self.button4.configure(font="-family {Poppins Regular} -size 10")
        self.button4.configure(borderwidth="0")
        self.button4.configure(text="""GENERATE TOE""")
        self.button4.configure(command=lambda:self.doc_generation())

        #Progress Bar
        self.pb = ttk.Progressbar(top, style="white.Horizontal.TProgressbar", orient='horizontal', mode='determinate', length=280)
        self.pb.place(relx=0.045, rely=0.82, width=818, height=20)
                
       
    #opens dialoge box and save the file name into the global variable from the path
    def select_file(self, variable_name):
       filetypes = (
        ('All Files', '*'),  
        ('Excel files', '*.xlsx'),
        ('Word doc', '*.docx')
        )

       filename = fd.askopenfile(
           mode = 'r',
           title = 'select a file',
           filetypes = filetypes)
    
       variable_name.set(filename.name)

    #open a dialoge box and save the path of the folder
    def select_dic(self,variable_name):
        filename = fd.askdirectory(title='Select a Folder')    
        variable_name.set(filename)
  
    '''

    Generates the word doc into the seleced folder
    Clears the varaibles after the doc is generated/when encountering an error

    '''
    def doc_generation(self, Event=None):
        folder_name = output_folder_name.get()
        excel_file = excel_name.get()
        template_file = template_name.get()

        os.chdir(folder_name)
        df = pd.read_excel(excel_file, engine='openpyxl')
        doc = DocxTemplate(template_file)
        total_doc = df.shape[0] #keep track of doc generated
        progress_by = 818/int(total_doc)

        try: #handle cases where user might not select correct file format or not select any file/folder before clicking the button
            for i, row in df.iterrows():
            #print(df['COPY PASTE "Prénom"'][i])
                context = {'name': df['COPY PASTE "Prénom"'][i].title()+' '+df['COPY PASTE "NOM"'][i].upper(),
                           'address': (df['COPY PASTE "Adresse"'][i]).title(),
                           'Start_Date': (df['START DATE'][i]).title().lower(),
                           'JOB_TITLE': (df['JOB TITLE'][i]),
                           'CBA_Level': (df['CBA Level'][i]),
                           'Comp': (df['ANNUAL COMPENSATION'][i]),
                           'SITE_NAME': (df['SITE NAME'][i]),
                           'DEPARTEMENT' : df['DEPARTEMENT'][i]
                           }
                doc.render(context)
                doc.save((df['COPY PASTE "Prénom"'][i].title()+' '+df['COPY PASTE "NOM"'][i].upper())+'.docx') 
                
                #updates the progress bar
                #idletask makes sure the bar progress do not happens in one swift flow
                self.pb['value'] += progress_by
                root.update_idletasks()
                

            print(i)
            messagebox.askokcancel("Result","TOE Generated Succefully for "+str(total_doc)+" Candidates")
            self.pb['value'] = 0
            excel_name.set("")
            template_name.set("")
            output_folder_name.set("")  

        except:
            messagebox.askokcancel("Error!","Oops! "+ str(sys.exc_info()[0]) + " occurred.\n\nPlease check the below points before proceeding\n\n1. Excel\n   * No additional column added or deleted\n   * Columns are named correctly\n   * All Data is updated correctly/not missing\n2. There is no change in the Template\n3. Files selected via buttons are in proper format\n   * Excel should be .xlsx format \n   * Template should be .docx format")
            self.pb['value'] = 0
            excel_name.set("")
            template_name.set("")
            output_folder_name.set("")                  

#exit handler          
def exitt():
    sure = messagebox.askyesno("Exit","Are you sure you want to exit?")
    if sure == True:
        root.destroy()

page1 = main_page(root)
root.protocol("WM_DELETE_WINDOW", exitt)
root.mainloop()