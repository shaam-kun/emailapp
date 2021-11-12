from tkinter import *
import tkinter
import os
from tkinter import filedialog
from tkinter import messagebox
import smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def open_file():
    EMAILAPP.root.filename = filedialog.askopenfilename(initialdir="/", title="Choose your file", filetype=(
        ("excel files", "*.xls,xlsx"), ("all Files", "*.")))
    return EMAILAPP.root.filename


def save_file_as():
    EMAILAPP.root.filename3 = filedialog.asksaveasfilename(initialdir="/", title="Save file as", filetypes=(
        ("excel files", "*.xls,xlsx"), ("all files", "*.*")))
    return EMAILAPP.root.filename3


def about_app():
    # var = StringVar()
    # label = Message(EMAILAPP.root, textvar=var, relief=SUNKEN, justify=CENTER, anchor=N, bg="black", fg="white")
    # var.set("\nThis app is designed to send emails!\n\n\nCreated by:\n\nHishaam")
    # label.pack(side=BOTTOM)
    messagebox.showinfo("ABOUT", "\nThis app is designed to send emails!\n\n\nCreated by:\n\nShaam")


def askdir():
    EMAILAPP.root.dir = filedialog.askdirectory()
    return EMAILAPP.root.dir


def save_file():
    EMAILAPP.root.filename2 = filedialog.asksaveasfile(initialdir=str(askdir()), title="Save file", filetypes=(
        ("excel files", "*.xls,xlsx"), ("all files", "*.*")))
    return EMAILAPP.root.filename2


class EMAILAPP:
    root = tkinter.Tk()
    name = " "
    srname = " "
    contact = " "
    email_add = " "

    # password = " "

    def __init__(self):
        self.filename = PhotoImage(file="c3.png").zoom(x=5, y=5)
        EMAILAPP.root.iconbitmap("images.ico")
        EMAILAPP.root.geometry("500x500")
        EMAILAPP.root.title("Email run.app")
        self.bkground = Label(EMAILAPP.root, image=self.filename)
        self.bkground.place(x=0, y=0, relwidth=1, relheight=1)
        self.menubar(EMAILAPP.root)
        Label(text="").pack()
        Button(text="Login", height="2", width="30", command=self.login).pack()
        Label(text="").pack()

    # TODO FIX NEW WINDOW BACKGROUND
    def New_window(self):
        newwindow = Toplevel(EMAILAPP.root)
        filename = PhotoImage(file="c3.png").zoom(x=5, y=5)
        newwindow.iconbitmap("images.ico")
        newwindow.geometry("500x500")
        newwindow.title("Email run.app")
        bkground = Label(newwindow, image=self.filename)
        bkground.place(x=0, y=0, relwidth=1, relheight=1)
        self.menubar(newwindow)
        # return EMAILAPP.menubar(newwindow)

    def menubar(self, root):
        menubar = tkinter.Menu(root)
        file_men = tkinter.Menu(menubar, tearoff=0)
        file_men.add_command(label="New", command=self.New_window)
        file_men.add_command(label="Open", command=open_file)
        file_men.add_command(label="Save", command=save_file)
        file_men.add_command(label="Save as...", command=save_file_as)
        file_men.add_separator()
        file_men.add_command(label="Exit", command=root.quit)
        menubar.add_cascade(label="File", menu=file_men)

        helpmenu = tkinter.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About...", command=about_app)
        menubar.add_cascade(label="Help", menu=helpmenu)
        root.config(menu=menubar)

    def register(self):  # TODO
        global register_screen
        register_screen = Toplevel(EMAILAPP.root)
        register_screen.iconbitmap("images.ico")
        register_screen.geometry("500x500")
        bkground = Label(register_screen, image=self.filename)
        bkground.place(x=0, y=0, relwidth=1, relheight=1)
        register_screen.title("ADD STUDENTS DETAILS")
        self.menubar(register_screen)
        EMAILAPP.name = StringVar()
        EMAILAPP.srname = StringVar()
        EMAILAPP.contact = StringVar()
        EMAILAPP.email_add = StringVar()
        # EMAILAPP.password = StringVar()
        global username_entry
        global surname_entry
        global contact_entry
        global email_entry
        Label(register_screen, text="Please enter your details below", bg="black", fg="white").pack()
        Label(register_screen, text="").pack()
        username_lable = Label(register_screen, text="Name * ")
        username_lable.pack()
        username_entry = Entry(register_screen, textvariable=EMAILAPP.name)
        username_entry.pack()
        surname = Label(register_screen, text="Surname * ")
        surname.pack()
        surname_entry = Entry(register_screen, textvariable=EMAILAPP.srname)
        surname_entry.pack()
        contact_lable = Label(register_screen, text="Contact * ")
        contact_lable.pack()
        contact_entry = Entry(register_screen, textvariable=EMAILAPP.contact)
        contact_entry.pack()
        email_lable = Label(register_screen, text="Email * ")
        email_lable.pack()
        email_entry = Entry(register_screen, textvariable=EMAILAPP.email_add)
        email_entry.pack()
        Label(register_screen, text="").pack()
        Button(register_screen, text="ADD STUDENT", width=10, height=1, bg="grey", fg="white",
               command=self.register_user).pack()

    def register_user(self):
        username_info = EMAILAPP.name.get()
        username_info = username_info.lower()
        email_info = EMAILAPP.email_add.get()
        file_name = "DATASTUDENTS.txt"
        file = open(file_name, "a+")
        file.write(username_info + " ")
        file.write(email_info + "\n")
        file.close()
        username_entry.delete(0, END)
        surname_entry.delete(0, END)
        contact_entry.delete(0, END)
        email_entry.delete(0, END)
        messagebox.showinfo("ADDED STUDENT", "SUCCESS")
        register_screen.destroy()

    # _______________________________________________________________________________________________________________________

    def login(self):
        global login_screen
        login_screen = Toplevel(EMAILAPP.root)
        login_screen.iconbitmap("images.ico")
        login_screen.geometry("500x500")
        bkground = Label(login_screen, image=self.filename)
        bkground.place(x=0, y=0, relwidth=1, relheight=1)
        login_screen.title("LOGIN")
        self.menubar(login_screen)
        global email_verify
        global password_verify
        email_verify = StringVar()
        # password_verify=StringVar()
        global email__loginentry
        global password__loginentry
        Label(login_screen, text="Please enter your details below to login", bg="black", fg="white").pack()
        Label(login_screen, text="").pack()
        Label(login_screen, text="EMAIL*").pack()
        email__loginentry = Entry(login_screen, textvariable=email_verify)
        email__loginentry.pack()
        # Label(login_screen,text="").pack()
        # password__loginentry = Entry(login_screen,textvariable=password_verify,show="*")
        # password__loginentry.pack()
        Label(login_screen, text="").pack()
        Button(login_screen, text="Login", width=10, height=1, bg="grey", fg="white", command=self.login_verify).pack()

    def login_verify(self):
        email1 = email_verify.get()
        #        password1=password_verify.get()
        email__loginentry.delete(0, END)
        # password__loginentry.delete(0,END)

        if email1 == "taliyahhifth1@gmail.com":
            # if password1 == "":
            messagebox.showinfo("LOGIN", "SUCCESS")
            self.login_suc()
            login_screen.destroy()
            # else:
            # self.password_not_recognised()
        else:
            self.user_not_found()

    #    def password_not_recgnised(self):
    #        global password_not_rec_screen
    #        password_not_rec_screen=Toplevel(login_screen)
    #        self.menubar(password_not_rec_screen)
    #        password_not_rec_screen.geometry("150x100")
    #        Label(password_not_rec_screen,text="INVALID PASSWORD").pack()
    #        Button(password_not_rec_screen,text="OK",command=self.del_pass_notrec).pack()

    def user_not_found(self):
        global user_nf_screen
        user_nf_screen = Toplevel(login_screen)
        self.menubar(user_nf_screen)
        user_nf_screen.title("OOPS")
        user_nf_screen.geometry("150x100")
        Label(user_nf_screen, text="INVALID EMAIL").pack()
        Button(user_nf_screen, text="OK", command=self.del_user_nf).pack()

    # def del_pass_notrec(self):
    #    password_not_rec_screen.destroy()
    def del_user_nf(self):
        user_nf_screen.destroy()

    def login_suc(self):
        Label(text=" ").pack()
        Button(text="ADD DETAILS", command=self.register).pack()
        Label(text="").pack()
        Button(text="SEND EMAIL", command=self.email_screen).pack()

    def add_mail(self):
        file2 = open("DATASTUDENTS.txt", "r")
        line2 = [k.split()[1] for k in file2]
        return line2

    def add_Students(self):
        file = open("DATASTUDENTS.txt", "r")
        line = [r.split()[0] for r in file]
        line = list(line)
        return line

    def send(self, emailtoreceive, file, password1):
        subject = "Ta'liyah Hifth School Student Progress Report"
        body = "This is an email with attachment sent from Madrasatut Ta'liyah Hafith School"
        sender_email = "taliyahhifth1@gmail.com"
        receiver_email = emailtoreceive
        password = password1

        # Create a multipart message and set headers
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject
        message["Bcc"] = receiver_email  # Recommended for mass emails

        # Add body to email
        message.attach(MIMEText(body, "plain"))

        filename = file  # "{fname}+{format1}".format(fname=file,format1=".xlsx") # In same directory as script

        # Open PDF file in binary mode
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )

        # Add attachment to message and convert message to string
        message.attach(part)
        text = message.as_string()

        # Log in to server using secure context and send email
        context = ssl.create_default_context()
        try:

            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:

                server.login(sender_email, password)
                server.sendmail(sender_email, receiver_email, text)
                messagebox.showinfo("EMAIL","SUCCESS")
                email_screen.destroy()
        except smtplib.SMTPServerDisconnected:
            messagebox.showinfo("EMAIL", "PASSWORD INVALID")
            email_screen.destroy()
        except smtplib.SMTPException as e:
            messagebox.showinfo("EMAIL","SOME ERROR HAS OCCURED:"+str(e))



    def process_send(self, email, file, password):
        self.send(email, file, password)

    def email_screen(self):
        global email_screen
        email_screen = Toplevel(EMAILAPP.root)
        email_screen.iconbitmap("images.ico")
        email_screen.geometry("500x500")
        bkground = Label(email_screen, image=self.filename)
        bkground.place(x=0, y=0, relwidth=1, relheight=1)
        email_screen.title("SEND EMAIL")
        self.menubar(email_screen)
        Label(email_screen, text="ENTER THE MAIL YOU WOULD LIKE TO SEND TO", font=("helvetica,10"), bg="grey").pack()
        global email_to_send
        global file_to_send
        global pass_word
        EMAILAPP.var1 =StringVar()
        EMAILAPP.var2 =StringVar()
        EMAILAPP.var3 =StringVar()
        Label(email_screen, text="").pack()
        email_to_send = Entry(email_screen, textvariable=EMAILAPP.var1).pack()
        Label(email_screen, text="ENTER THE FileNAME with extenstion", font=("helvetica,10"), bg="grey").pack()
        Label(email_screen, text=" ").pack()
        file_to_send = Entry(email_screen, textvariable=EMAILAPP.var2).pack()
        Label(email_screen, text="ENTER THE PASSWORD", font=("helvetica,10"), bg="grey").pack()
        Label(email_screen, text=" ").pack()
        pass_word = Entry(email_screen, textvariable=EMAILAPP.var3, show="*").pack()
        Label(email_screen, text=" ").pack()

        Button(email_screen, text="SEND", font=("helvetica,10"), bg="grey", fg="white",
               command=self.run).pack()
        # file = open("DATASTUDENTS.txt", "r")
        # line = [r.split()[0] for r in file]
        # line2 = [k.split()[1] for k in file2]
        # Label(email_screen, text="Also choose the name of the student", bg="black", fg="white").pack()
        # Label(email_screen, text="Names of students", bg="black", fg="white").pack()
        # for i in range(len(line)):
        # Checkbutton(email_screen,text=line[i],onvalue = 1,offvalue =2).pack(side="left")
        # Label(email_screen, text=" ").pack(side="right")
        # Label(email_screen, text=" ").pack(side="right")
        # Label(email_screen, text=" ").pack(side="right")
        # for k in range(len(line2)):
        # Checkbutton(email_screen, text=line2[k], onvalue=1, offvalue=2).pack(side="right")
        # Label(email_screen, text=" ").pack(side="left")
        # taliyahhifth1@gmail.com

    def run(self):
        em = EMAILAPP.var1.get()
        fil =EMAILAPP.var2.get()
        pas = EMAILAPP.var3.get()
        self.process_send(em, fil, pas)


class Student:
    name = ""
    srname = ""
    contact = ""
    email_add = ""

    def __init__(self, name, email_ad):
        self.name = name
        self.email_add = email_ad

    """SETTERS"""

    def setname(self, name):
        self.name = name

    def set_email(self, email):
        self.email_add = email

    """GETTERS"""

    def getname(self):
        return self.name

    def getemail(self):
        return self.email_add

    def add_Students(self):
        file = open("DATASTUDENTS.txt", "r")
        line = [r.split()[0] for r in file]
        return line

    def add_mail(self):
        file2 = open("DATASTUDENTS.txt", "r")
        line2 = [k.split()[1] for k in file2]
        return line2


app = EMAILAPP()
app.root.mainloop()
app2 = Student("", "")
app2.add_Students()
