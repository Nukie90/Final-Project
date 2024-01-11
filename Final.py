import csv
import webbrowser
from abc import ABC, abstractmethod
import sys
import subprocess

try:
    from tkinter import * # python 3
    import tkinter.messagebox as mb
except ImportError:
    from TKinter import * # python 2
    from Tkinter import messagebox
    
try:
    import customtkinter as ct
except ImportError:
    print("> 'customrkinter' module is missing!\n" +
        "Trying to install required module: customtkinter")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "customtkinter"])
    print()
finally:
    import customtkinter as ct
    
try:
    from fpdf import FPDF
except ImportError:
    print("> 'customrkinter' module is missing!\n" +
        "Trying to install required module: PyPDF2")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "fpdf"])
    print()
finally:
    from fpdf import FPDF
    
ct.set_default_color_theme('dark-blue')

class Window(ABC):
    def __init__(self,win = None):
        self.win = win
        self.win = ct.CTk()
        self.win.geometry('400x400+550+200')
        self.win.title('Contact Book')
        self.win.resizable(width=False, height=False)
        self.label = ct.CTkLabel(self.win, text="CONTACTBOOK", font=('arial', 20, 'bold')).place(anchor='n',relx=0.5, rely= 0.05)
        
        self.name = StringVar()
        self.num = StringVar()
        self.email = StringVar()
        self.instagram = StringVar()
        self.facebook = StringVar()
        self.wish = StringVar()
        
        labelFrame = ct.CTkFrame(master=self.win)
        labelFrame.place(relx=0.1,rely=0.17,relwidth=0.8,relheight=0.42)
        
        self.scroll = ct.CTkScrollbar(master= labelFrame, )
        self.scroll.pack(side=RIGHT, fill=Y)
        self.list = Listbox(master=labelFrame, yscrollcommand=self.scroll.set, width=200, font=('Arial', 16))

        try:
            with open('database.csv', newline='') as f:
                self.reader = csv.reader(f)
                self.data = list(self.reader)
                self.data.sort()

        except:
            with open('database.csv', 'w') as file:
                self.writer = csv.writer(file)
                self.data = []
                
            mb.showinfo(message="App has just created new file for you \nEnjoy!")
                
        self.list.pack(side=LEFT, fill=BOTH)
        
        for i in self.data:
            self.list.insert(END, i[0])
        
        button1 = ct.CTkButton(self.win,text="ADD", command = self.add, width=20, hover_color="blue")
        button1.place(relx=0.1, rely=0.65, anchor='n')
        button2 = ct.CTkButton(self.win,text="SAVE",command = self.save, width=20, hover_color="blue")
        button2.place(relx=0.1, rely=0.8, anchor='n')
        button3 = ct.CTkButton(self.win,text="VIEW",command = self.view_select, width=20, hover_color="blue")
        button3.place(relx=0.5, rely=0.65, anchor='n')
        button4 = ct.CTkButton(self.win,text="DELETE",command = self.delete, width=20, hover_color="blue")
        button4.place(relx=0.9, rely=0.65, anchor='n')
        button5 = ct.CTkButton(self.win,text="CLEAR",command = self.clear, width=25, hover_color="blue")
        button5.place(relx=0.9, rely=0.8, anchor='n')
        menu1 = ct.CTkOptionMenu(self.win, values=["System", "Dark", "Light"], command = self.appearance,)
        menu1.place(relx=0.5, rely=0.8, anchor='n')
        exitbutton = ct.CTkButton(self.win,text="QUIT", command = self.quit, width=20, hover_color="red")
        exitbutton.place(relx=0.5, rely=0.9, anchor='n')
        
        self.win.mainloop()
        
    @abstractmethod
    def add(self):
        pass
    
    @abstractmethod
    def save(self):
        pass
    
    @abstractmethod
    def view_select(self):
        pass
    
    @abstractmethod
    def delete(self):
        pass
    
    @abstractmethod
    def clear(self):
        pass
    
    @abstractmethod
    def appearance(self):
        pass

    @abstractmethod
    def quit(self):
        pass 

class insidecommand(Window):
    def enter(self):
        try:
            # When user completely enter all infomation
            if self.name.get()!= '' and self.num.get() != '' and self.email.get() != '' and self.instagram.get() != '' and self.facebook.get() != '':
                self.data.append([self.name.get(),self.num.get(),self.email.get(),self.instagram.get(),self.facebook.get()])
                self.list.insert(END, self.name.get())
                self.name.set('')
                self.num.set('')
                self.email.set('')
                self.instagram.set('')
                self.facebook.set('')
                self.add.destroy()
            # When user forgot or didn't enter '-' in entry.
            elif self.name.get() == '' or self.num.get() == '' or self.email.get() == '' or self.instagram.get() == '' or self.facebook.get() == '':
                raise ValueError

        except ValueError:
            mb.showerror(title="ValueError", message='Please add - to the empty entry')
            
    def to_pdf(self):
        NAME, TELE, MAIL, IG, FB = self.data[self.select()]
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size = 24)
        
        pdf.cell(200, 10, txt = f"Name: {NAME}", ln = 1, align= 'W')
        pdf.cell(200, 10, txt = f"Name: {TELE}", ln = 2, align= 'W')
        pdf.cell(200, 10, txt = f"Name: {MAIL}", ln = 3, align= 'W')
        pdf.cell(200, 10, txt = f"Name: {IG}", ln = 4, align= 'W')
        pdf.cell(200, 10, txt = f"Name: {FB}", ln = 5, align= 'W')
            
        pdf.output(f'{NAME}.pdf')
        mb.showinfo(message="PDF file has been created")
class Contact(insidecommand,Window):
    def __init__(self, win= None):
        super().__init__(win)   

    def select(self):
        try:
            return int(self.list.curselection()[0])
        except IndexError:
            mb.showerror(title='INDEX ERROR',message='Please select a contact')
    #Enter string
    def add(self):
        self.add = ct.CTkToplevel(self.win)
        self.add.title("ADD")
        self.add.geometry('270x200+550+200')
        self.add.resizable(width=False, height=False)
        
        self.name.set('')
        self.num.set('')
        self.email.set('')
        self.instagram.set('')
        self.facebook.set('')
       
        label1 = ct.CTkLabel(self.add, text="Name: ")
        label1.grid(row=0, column=0)
        entry1 = ct.CTkEntry(self.add, textvariable=self.name,width=200, height=20)
        entry1.grid(row=0, column=1)
        
        label2 = ct.CTkLabel(self.add, text="Number: ")
        label2.grid(row=1, column=0)
        entry2 = ct.CTkEntry(self.add, textvariable=self.num,width=200, height=20)
        entry2.grid(row=1, column=1)
        
        label3 = ct.CTkLabel(self.add, text="Email: ")
        label3.grid(row=2, column=0)
        entry3 = ct.CTkEntry(self.add, textvariable=self.email, width=200, height=20)
        entry3.grid(row=2, column=1)
        
        label4 = ct.CTkLabel(self.add, text="Instagram: ")
        label4.grid(row=3, column=0)
        entry4 = ct.CTkEntry(self.add, textvariable=self.instagram, width=200, height=20)
        entry4.grid(row=3, column=1)
        
        label5 = ct.CTkLabel(self.add, text="Facebook: ")
        label5.grid(row=4, column=0)
        entry5 = ct.CTkEntry(self.add, textvariable=self.facebook, width=200, height=20)
        entry5.grid(row=4, column=1)
        
        button1 = ct.CTkButton(self.add, text="Enter",hover_color="blue", command=self.enter)
        button1.grid(row=5,columnspan=2,column=1, sticky="NSEW")
        
    #save the list to csv file
    def save(self):
        with open("database.csv", 'w',newline="") as f:
            self.write = csv.writer(f,)
            self.write.writerows(self.data)
        mb.showinfo(message="All contacts have been saved")

    
    def view_select(self):
        NAME, TELE, MAIL, IG, FB = self.data[self.select()]
        self.view_win = ct.CTkToplevel(self.win)
        self.view_win.title(NAME)
        self.view_win.geometry('350x300+550+200')
        self.view_win.resizable(height=False, width=FALSE)
        
        ct.CTkLabel(self.view_win, text= f'Name: ', font=('arial', 15),).place(relx=0.1,rely=0)
        ct.CTkLabel(self.view_win, text= f'Telephone: ', font=('arial', 15)).place(relx=0.1,rely=0.12)
        ct.CTkLabel(self.view_win, text= f'E-Mail: ', font=('arial', 15)).place(relx=0.1,rely=0.24)
        ct.CTkLabel(self.view_win, text= f'Instagram: ', font=('arial', 15)).place(relx=0.1,rely=0.36)
        ct.CTkLabel(self.view_win, text= f'Facebook: ', font=('arial', 15)).place(relx=0.1,rely=0.48)
        
        ct.CTkEntry(self.view_win, textvariable=self.name,width=200, height=20).place(relx=0.35,rely=0)
        ct.CTkEntry(self.view_win, textvariable=self.num,width=200, height=20).place(relx=0.35,rely=0.12)
        ct.CTkEntry(self.view_win, textvariable=self.email,width=200, height=20).place(relx=0.35,rely=0.24)
        ct.CTkEntry(self.view_win, textvariable=self.instagram,width=200, height=20).place(relx=0.35,rely=0.36)
        ct.CTkEntry(self.view_win, textvariable=self.facebook,width=200, height=20).place(relx=0.35,rely=0.48)
        
        self.name.set(NAME)
        self.num.set(TELE)
        self.email.set(MAIL)
        self.instagram.set(IG)
        self.facebook.set(FB)
        
        ct.CTkButton(self.view_win, text='PDF', command=self.to_pdf).place(relx=0.32,rely=0.7)
        ct.CTkButton(self.view_win, text='To Instagram', command=lambda: webbrowser.open(f'https://www.instagram.com/{IG}/')).place(relx=0.05,rely=0.8)
        ct.CTkButton(self.view_win, text='To Facebook', command=lambda: webbrowser.open(f'https://www.facebook.com/{FB}')).place(relx=0.565,rely=0.8)

    def delete(self):
        try:
        # get selected line index
            del self.data[self.select()]
            index = self.list.curselection()[0]
            self.list.delete(index)
        
        except TypeError:
            pass
        
    def clear(self):
        self.list.delete(0,END)
        with open('database.csv', 'w') as file:
            file.write('')

    def appearance(self, new_appearance_mode):
        ct.set_appearance_mode(new_appearance_mode)
        
    def quit(self):
        self.win.destroy()
        
if __name__ == "__main__":
    Contact()