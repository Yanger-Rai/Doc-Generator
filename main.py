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
from datetime import date
import warnings
import re
import time
from datetime import datetime
import datetime as DT


#initiating UI
root = Tk()
root.title("AU TOE Tool")

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
        top.title("AU TOE Tool")

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
        self.button1.configure(command=lambda:self.select_excel(excel_name))

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
        self.button2.configure(command=lambda:self.select_template(template_name))

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
    def select_excel(self, variable_name):
       filetypes = [('Excel files', '*.csv')]

       filename = fd.askopenfile(
           mode = 'r',
           title = 'select a file',
           filetypes = filetypes)
    
       variable_name.set(filename.name)
    
    def select_template(self, variable_name):
       filetypes = [('Word doc', '*.docx')]

       filename = fd.askopenfile(
           mode = 'r',
           title = 'select a file',
           filetypes = filetypes)
    
       variable_name.set(filename.name)

    #open a dialoge box and save the path of the folder
    def select_dic(self,variable_name):
        filename = fd.askdirectory(title='Select a Folder')    
        variable_name.set(filename)
  
    #Validates if a column has all data in interger types and throws error is condition not met
    def interger_validation(self, data_frame, column_name):
        matched_digit = []

        for data in data_frame[column_name]:
            x = re.findall("[a-zA-Z]", str(data)) #will look for character alphabetically between a and z, lower case OR upper case 
            if x:
                matched_digit.append(1)
            else:
                matched_digit
    
        if 1 in matched_digit:   
            messagebox.askretrycancel("ERROR!", "please check the data in "+ str(column_name) +" column and try again!")
            sys.exit()
        else:
            return
    #Validates if a column has all data in string types and throws error is condition not met   
    def string_validation(self, data_frame, column_name):
        matched_digit = []

        for data in data_frame[column_name]:
            x = re.findall("\d", str(data)) #will look for a digit 
            if x:
                matched_digit.append(1)
            else:
                matched_digit
    
        if 1 in matched_digit:   
            messagebox.askretrycancel("ERROR!", "please check the data in "+ str(column_name) +" column and try again!")
            sys.exit()
        else:
            return


    '''
    Generates the word doc into the seleced folder
    Clears the varaibles after the doc is generated/when encountering an error
    '''
    def doc_generation(self, Event=None):

        confirmation = messagebox.askyesno ("confirmation", "Proceed to generate TOE?")
        if confirmation is False:
            excel_name.set("")
            template_name.set("")
            output_folder_name.set("")   
            return        

        folder_name = output_folder_name.get()
        excel_file = excel_name.get()
        template_file = template_name.get()

        os.chdir(folder_name)
        #df = pd.read_excel(excel_file, engine='openpyxl')
        df = pd.read_csv(excel_file)
        doc = DocxTemplate(template_file)
        total_doc = df.shape[0] #keep track of doc generated
        progress_by = 818/int(total_doc) #divide helps in visually matching the total number of file generation
        st = time.process_time()

        #Data validation        
        #self.interger_validation(df,'ANNUAL COMPENSATION')
        #self.string_validation(df, 'COPY PASTE "Prénom"')
        #self.string_validation(df, 'COPY PASTE "NOM"')
        #self.string_validation(df, 'DEPARTEMENT')

        def integer_converter_formatter(number):
            # EU format for thousand and lakh with decimal aaa.aaa,aa
            pattern_EU_with_dec = "^\d+\.\d+\,\d{2}$"
            # EU format for thousand and lakh without decimal aaa.aaa
            pattern_EU_without_dec = "^\d+\.\d{3}$"
            # EU format for hundred with decimal aaa,aa
            pattern_EU_hundred = "^\d{3}\,\d{2}$"

            x = re.search(pattern_EU_with_dec, number)
            y = re.search(pattern_EU_without_dec, number)
            z = re.search(pattern_EU_hundred, number)
            
            if isinstance(number, str):
                # if the number is not string, return as it is
                # if the number is string
                # Check if it in EUR or US format
                # And convert it to aaaa.aa format
                # this format is suitable for float/int convert using builtin method    
                new_number = ""
                if x is not None:
                    new_number = number.replace(".","").replace(",",".")
                    return (new_number)        
                elif y is not None:
                    new_number = number.replace(".","")
                    return (new_number)            
                elif z is not None:
                    new_number = number.replace(",",".")
                    return (new_number)        
                return (re.sub("[,$]","",number))      
            return number


        try: #handle cases where user might not select correct file format or not select any file/folder before clicking the button
            for i, row in df.iterrows():
                startdate = df['4 - Offer Letter Components : OfferDatePositionStart'][i]
                currentdate = date.today()

                context = {'CANDIDATE_NAME': df['Person : First Name'][i]+" "+df['Person : Last Name'][i],
                           'ADDRESS': (str(df['Home: Addresses : Street'][i]) +"\n"+ str(df['Home: Addresses : Street 2'][i]) +"\n"+ str(df['Home: Addresses : Zip/Postal Code'][i]) + " " + str(df['Home: Addresses : State/Province (Full Name)'][i]) +"\n"+ str(df['Home: Addresses : Country (Full Name)'][i])),
                           'NAME': df['Person : First Name'][i],
                           'START_DATE' : datetime.strptime(startdate, "%m/%d/%Y").strftime("%d %B %Y"),
                           'TITLE': df['Job : External Job Title'][i],
                           'LOCATION': df['Company/Location (search) : Street'][i],
                           'MANAGER_TITLE': df['Hiring Manager : Full Name: First Last'][i],
                           'WAGE': df['4 - Offer Letter Components : OfferBaseSalary'][i],
                           'TOTAL_BONUS': int(float(integer_converter_formatter(df['4 - Offer Letter Components : OfferSplitBonusYear2'][i]))) + int(float(integer_converter_formatter(df['4 - Offer Letter Components : OfferSplitBonusYear1'][i]))),
                           'BONUS_1': int(float(integer_converter_formatter(df['4 - Offer Letter Components : OfferSplitBonusYear2'][i]))),
                           'BONUS_2': int(float(integer_converter_formatter(df['4 - Offer Letter Components : OfferSplitBonusYear1'][i]))),
                           'DATE' : (currentdate + DT.timedelta(days=7)) #7 days after the current date
                           }
                                
                doc.render(context)
                doc.save((df['Person : First Name'][i]+" " +df['Person : Last Name'][i])+'.docx') 
                    
                #updates the progress bar
                #idletask makes sure the bar progress do not happens in one swift flow
                self.pb['value'] += progress_by
                root.update_idletasks()
                

            print(i)
            messagebox.showinfo("Result","TOE Generated Succefully for "+str(total_doc)+" Candidates\n\nMake sure to give a 4eye check on the documents generated")
            self.pb['value'] = 0
            excel_name.set("")
            template_name.set("")
            output_folder_name.set("")  

        except:
            messagebox.showerror("Error!","Oops! "+ str(sys.exc_info()[0]) + " occurred.\n\nPlease check the below points before proceeding\n\n1. Excel\n   * No additional column added or deleted\n   * Columns are named correctly\n   * All Data is updated correctly/not missing\n2. There is no change in the Template\n3. Files selected via buttons are in proper format\n   * Excel should be .xlsx format \n   * Template should be .docx format")
            self.pb['value'] = 0
            excel_name.set("")
            template_name.set("")
            output_folder_name.set("")

        et = time.process_time()
        res = et - st
        print('CPU Execution time:', res, 'seconds')                  

#exit handler          
def exitt():
    sure = messagebox.askyesno("Exit","Are you sure you want to exit?")
    if sure == True:
        root.destroy()

page1 = main_page(root)
root.protocol("WM_DELETE_WINDOW", exitt)
root.mainloop()