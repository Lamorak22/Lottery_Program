#KOA lottery sheet
#For GUI
from tkinter import *
import webbrowser
import tkinter.messagebox
import subprocess

#For Excel
from openpyxl import Workbook, load_workbook

#For obtaining current time
import time

class KOA_Lottery: #Use master for Tk functions and commands
    def __init__(self, master):
        self.master = master
        master.title('KOA Lottery Sheet')

        ##### Display the GUI in the middle of the screen with the appropriate dimensions
        # w = 1000
        # h = 700
        # sw = master.winfo_screenwidth()
        # sh = master.winfo_screenheight()
        # x = (sw - w)/2
        # y = (sh - h)/2
        # master.geometry('%dx%d+%d+%d' % (w,h,x,y))  
        

        #####
        f1 = Frame(master)
        f1.pack(side=LEFT)
        f2 = Frame(master)
        f2.pack(side=RIGHT)

        ##### Get current date and create a string that symbolizes the =DATE formula in Excel
        self.local_time = time.localtime(time.time())
        self.curr_time =  str(self.local_time[1]) + "/" + str(self.local_time[2]) + "/" + str(self.local_time[0]) 

        ##### Making the file menu at the top of the application
        menu = Menu(master)
        master.config(menu=menu)
        menu.add_cascade(label='About', command=self.aboutWindow)

        ##### Initialize inventory
        self.inventory={}
        self.loadInventory()

        ##### Make 25 Labels
        for i in range(1,25):
           lbltxt = "Lottery #" + str(i) + ":"
           l1 = Label(f1, text=lbltxt, font="bold")
           l1.grid(row=i, sticky=W)

        ##### Make 25 Entries and store in dictionary
        self.d={}            
        for x in range(1,25):
            self.d[f'e{x}'] = Entry(f1, bg="yellow")
            self.d[f'e{x}'].grid(row=x, column=1)
        
        
        ##### Display KOA image in top right
        photo = PhotoImage(file="C:\\Users\\Daniel\\Pictures\\koa-logo.png")
        w = Label(f2, image=photo)
        w.photo = photo
        w.grid(row=0,column=0,sticky=N+E)

        ##### Buttons for when a lottery is out of stock or out
        out_btn = Button(f2, height=2, bg="green", text="Click this button if a lottery is out", font="bold", command=self.lotteryOut)
        out_btn.grid(row=1,column=0,columnspan=2,sticky=W+E+N+S)
        stock_btn = Button(f2, height=2, bg="Green", text="Click this button to restock a lottery", font="bold", command=self.restockPopup)
        stock_btn.grid(row=2,column=0,columnspan=2,sticky=W+E+N+S)

        ##### Assigning function to buttons  
        submitButton = Button(f2, text="Submit", bg="yellow", height=2, font="bold", command=lambda:[self.export_to_excel(), self.update_amt_sold(), master.quit()])
        submitButton.grid(row=3,column=0,columnspan=2,sticky=W+E+N+S) 
        quitButton = Button(f2, text="Quit", bg="red", height=2, font="bold", command=master.quit)
        quitButton.grid(row=4,column=0,columnspan=2,sticky=W+E+N+S)

    def popUpConstructor(self, popup, w, h):
        ##### Display the new window in the middle of the screen with the appropriate dimensions
        sw = popup.winfo_screenwidth()
        sh = popup.winfo_screenheight()
        x = (sw - w)/2
        y = (sh - h)/2
        popup.geometry('%dx%d+%d+%d' % (w,h,x,y))
    
    def aboutWindow(self):
        popup = Tk()
        self.popUpConstructor(popup, 400, 200)
        popup.title("About")

        ##### Display appropriate info and contact information
        font = ("Helvetica", "15", "bold")
        msg = "Created By: Daniel Eberhart\nEmail: danieleberhart14@gmail.com"
        label = Label(popup, text=msg, width = 120, height=10, bg="yellow", font=font)
        label.pack()

        ##### Create exit button
        b1 = Button(popup, text="Exit", bg="yellow", width=10, command=popup.destroy)
        b1.pack()
        popup.mainloop()

    def export_to_excel(self):
        ##### Initialize variables
        row_num = 1
        col_num = 2
        date_found = False

        ##### Load workbook and open selected sheet from said workbook
        wb = load_workbook("KOA-Lottery-Excel.xlsx")
        ws = wb["Data"]

        ##### Get date
        while date_found == False:
            if ws.cell(row=row_num, column=1).value == str(self.curr_time):
                #print("Date Successfully found at row " + str(row_num))
                date_found = True
            else:
                row_num += 1
    
        ##### Writing to the cells
        for x in range(1,25):
            temp_num = self.d[f'e{x}'].get()
            ws.cell(row=row_num, column=col_num, value=int(temp_num))
            col_num += 1
        ##### Save Workbook    
        wb.save("KOA-Lottery-Excel.xlsx")

    def update_amt_sold(self):
        ##### Initialize variables
        col_num = 2
        row_num = 1
        date_found = False

        ##### Load workbook and open selected sheets from said workbook
        wb = load_workbook("KOA-Lottery-Excel.xlsx")
        ws1 = wb["Sold"]
        ws2 = wb["Data"]
        

        ##### Get date
        while date_found == False:
            if ws1.cell(row=row_num, column=1).value == str(self.curr_time):
                #print("Date Successfully found at row " + str(row_num))
                date_found = True
            else:
                row_num += 1

        ##### Assign sold data and calculate total sold
        total_sold = 0
        for i in range(1,25):
            minuend = ws2.cell(row=row_num, column=col_num).value
            subtrahend = ws2.cell(row=row_num-1, column=col_num).value
            difference = int(minuend) - int(subtrahend)
            total_sold += difference
            ws1.cell(row=row_num, column=col_num, value=int(difference))
            col_num += 1

        ##### Assign the total sold for today's date    
        ws1.cell(row=row_num, column=col_num, value=int(total_sold))

        ##### Save Workbook
        wb.save("KOA-Lottery-Excel.xlsx")

    def loadInventory(self):
        row_num = 2
        wb = load_workbook("KOA-Lottery-Excel.xlsx")
        ws = wb["Inventory"] 
        ##### Load inventory
        for x in range(1,25):
            self.inventory[f'inv{x}'] = ws.cell(row=row_num, column=2).value
            row_num += 1
        wb.save("KOA-Lottery-Excel.xlsx")

    def restockPopup(self):
        popup = Tk()
        self.popUpConstructor(popup, 400, 275)
        popup.title("Restock")
        msg = "Please click the Lottery that needs to be restocked."
        label1 = Label(popup, text=msg)
        label1.grid(row=0,column=0,columnspan=3) 

        label2 = Label(popup, text="Lottery # to restock (1-24): ")
        label2.grid(row=1,column=0, sticky=W)
        self.restock_num = Entry(popup, width=5)
        self.restock_num.grid(row=1, column=1)

        label3 = Label(popup, text="How many does the new lottery set have: ")
        label3.grid(row=2,column=0, sticky=W)
        self.restock_amt = Entry(popup, width=5)
        self.restock_amt.grid(row=2, column=1)

        btn_1 = Button(popup, text="Submit", command=lambda:[self.restockInventory(), popup.destroy()])
        btn_1.grid(row=3, rowspan=2, column=0, columnspan=3)

    def restockInventory(self):
        wb = load_workbook("KOA-Lottery-Excel.xlsx")
        ws = wb["Inventory"]
        lottery_num = self.restock_num.get()
        new_inventory_value = self.restock_amt.get()
        if int(lottery_num) <= 24 and int(lottery_num) >= 1:
            ws.cell(row=int(lottery_num), column=2, value=int(new_inventory_value))
        else:
            self.restockPopup()
        wb.save("KOA-Lottery-Excel.xlsx")

    def lotteryOut(self):
        popup = Tk()
        self.popUpConstructor(popup, 400, 275)
        popup.title("Restock")
        msg = "Please indicate which lottery # is out."
        label1 = Label(popup, text=msg)
        label1.grid(row=0,column=0,columnspan=3) 

        label2 = Label(popup, text="Lottery # that is out (1-24): ")
        label2.grid(row=1,column=0, sticky=W)
        self.restock_num = Entry(popup, width=5)
        self.restock_num.grid(row=1, column=1)



root = Tk()
mygui = KOA_Lottery(root)
root.mainloop()
