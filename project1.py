#JAYANTHI SUBRAHMANYA SAI
from tkinter import *
from tkinter import ttk
import  tkinter.messagebox
from tkinter.filedialog import askopenfilename
import pandas as pd
import imgkit
import os
import smtplib
import xlrd
from email.message import EmailMessage
import matplotlib.pyplot as plotter
class Main:
    def __init__(self,root):
        root = root
        root.title('GradeMailZ')
        root.configure(bg = 'Gray3')
        theLabel = Label(root,text = "WELCOME ADMIN!!",bg = "gray",fg = "black",font=("Comic Sans MS","16"))
        theLabel.pack(fill = X)
        frame = Frame(root,width=300,height=250)
        topFrame = Frame(root) 
        topFrame.pack()
        bottomFrame = Frame(root)
        bottomFrame.pack(side = BOTTOM) #Implicit Position
        button1 = Button(topFrame, text = "Import Results File",command = self.get_Results)
        button2 = Button(topFrame, text = "Import MailID(s) File",command = self.get_mailIDs)
        button3 = Button(topFrame, text = "DECLARE RESULTS!!!",command = self.send_All) #"command" used to bind 1 func with layout 1 widget
        button1.pack(side=LEFT)
        button2.pack(side = RIGHT)
        button3.pack(side =LEFT)
        frame.pack()
        root.focus_force() #making this window to work in foreground w/o being passive when it was started.....
        root.mainloop()
    def get_Results(self):
        Tk().withdraw()
        cols=['htno','code','sname','nothing','grade','credits']
        csv_file = askopenfilename()
        #if path.endswith('.csv'): Warning
        global result_Data
        result_Data = pd.read_csv(csv_file,names = cols,skipinitialspace = True)
        print("Results FILE")
    def get_mailIDs(self):
        Tk().withdraw()
        xl_file = askopenfilename()
        cols = ['Id','Mail']
        global mail_Ids
        mail_Ids = pd.read_excel(xl_file,names =cols)
        print("MailIDS file")
    def extract_StdResults(self,ht_no):
        #if path.endswith('.csv'):
        cols=['Ht.No','Code','Subject','nothing','Grade','Credits ']
        std_df = result_Data[result_Data.htno == ht_no]
        del std_df['nothing']
        std_df = std_df.rename(columns={'htno': 'HT.Number', 'code': 'Code','grade':'Grade','credits':'Credits','sname':'Subject/Laboratory'})
        std_df.set_index('HT.Number',inplace=True)
        return(std_df)
    def convert_to_HTML(self,std_df,std_id):
        data = std_df
        css = """
        <style type=\"text/css\">
        table {
        color: #333;
        font-family: Helvetica, Arial, sans-serif;
        width: 640px;
        border-collapse:
        collapse; 
        border-spacing: 0;
        }
        td, th {
        border: 1px solid transparent; /* No more visible border */
        height: 30px;
        }
        th {
        background: #DFDFDF; /* Darken header a bit */
        font-weight: bold;
        }
        td {
        background: #FAFAFA;
        text-align: center;
        }
        table tr:nth-child(odd) td{
        background-color: white;
        }
        </style>
        """
        std_id = std_id+".html"
        text_file = open(std_id,"a")
        text_file.write(css)
        text_file.write(data.to_html())
    def get_Cgpa(self,std_df):
        cgpa = 0
        marks_dict = {'O':10, 'S':9, 'A':8, 'B':7,'C':6,'D':5,'F':0,'ABSENT':0}
        std_marks = list(std_df.grade)
        std_credits = list(std_df.credits)
        for i in range(len(std_credits)):
            std_credits[i] = int(std_credits)
        for i in range(len(std_credits)):
            cgpa = cgpa+(std_marks[i]*std_credits[i])
        cgpa = cgpa/(std_df.shape[0])
        return cgpa
    def send_Mail(self,receiver,std_id):
        msg = EmailMessage()
        msg['subject'] = "Grades"
        mailadd = "js1theaspirant@gmail.com"
        msg['from'] = mailadd
        msg['to'] =  receiver
        msg.set_content("JNTUK")
        filename  = os.getcwd()+"\\"+std_id+".html"
        html_msg = open(filename).read()
        msg.add_alternative(html_msg,subtype = 'html')
        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
            smtp.login(mailadd,"jsnjssjuk")
            smtp.send_message(msg)
    def send_All(self):
        l1 = []
        l2 = []
        global mail_Ids
        for i in mail_Ids.Id:
            l1.append(i)
        for i in mail_Ids.Mail:
            l2.append(i)
        dict_1 = dict(zip(l1,l2))
        for i,j in dict_1.items():
            self.convert_to_HTML(self.extract_StdResults(i),i)
            #self.send_Mail(j,i)
        print("SUCCESS")

class Authentication:
    user = 'admin'
    passw = 'admin'
    def __init__(self,root):

        self.root = root
        self.root.title('USER AUTHENTICATION')

        '''Make Window 10X10'''

        rows = 0
        while rows<10:
            self.root.rowconfigure(rows, weight=1)
            self.root.columnconfigure(rows, weight=1)
            rows+=1

        '''Username and Password'''

        frame = LabelFrame(self.root, text='Login')
        frame.grid(row = 1,column = 1,columnspan=10,rowspan=10)
        #JAYANTHI SUBRAHMANYA SAI
        Label(frame, text = ' Usename ').grid(row = 2, column = 1, sticky = W)
        self.username = Entry(frame)
        self.username.grid(row = 2,column = 2)

        Label(frame, text = ' Password ').grid(row = 5, column = 1, sticky = W)
        self.password = Entry(frame, show='*')
        self.password.grid(row = 5, column = 2)

        # Button

        ttk.Button(frame, text = 'LOGIN',command = self.login_user).grid(row=7,column=2)
        #ttk.Button(frame, text='FaceID(Beta)',command = self.face_unlock).grid(row=8, column=2)

        '''Message Display'''
        self.message = Label(text = '',fg = 'Red')
        self.message.grid(row=9,column=6)


    def login_user(self):

        '''Check username and password entered are correct'''
        if self.username.get() == self.user and self.password.get() == self.passw:
            

            #Destroy current window
            root.destroy()
            
            #Open new window
            newroot = Tk()
            application = Main(newroot)
            
            newroot.mainloop()



        else:

            '''Prompt user that either id or password is wrong'''
            self.message['text'] = 'Username or Password incorrect. Try again!'





if __name__ == '__main__':
    global result_Data
    global mail_Ids
    root = Tk()
    root.geometry('425x185+700+300')
    application = Authentication(root)
    root.mainloop()


