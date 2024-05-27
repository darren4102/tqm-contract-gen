from datetime import datetime
import sys
import tkinter as tk
from tkinter import StringVar, messagebox, font, filedialog
import ttkbootstrap as tb
from docx import Document
from gen import CreateContract
import os

class MyGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("600x850")
        self.root.resizable(width=False, height=False)
        self.root.title("Contract Creator")
        digitFunc = self.root.register(self.validateNumber)
        inputFunc = self.root.register(self.validateInput)

        self.label = tk.Label(self.root, text="Create New Contract", font=font.Font(family="Calibri", size=24, weight="bold"))
        self.label.pack(padx=10, pady=10)

        self.entryFrame = tk.Frame(self.root)
        
        for i in range(3):
            self.entryFrame.columnconfigure(i, weight=1)

        # Title
        self.titleLabel = tb.Label(self.entryFrame, text="Title", font=('Calibri', 16))
        self.titleEntry = tb.Entry(self.entryFrame, font=('Calibri', 16), validate="focus", validatecommand=(inputFunc, '%P'))
        self.titleLabel.grid(row=0, column=0, sticky="w", padx=(5, 5))
        self.titleEntry.grid(row=1, column=0, sticky="ew", padx=(5, 5), pady=(0, 5))
        # First name
        self.firstnameLabel = tb.Label(self.entryFrame, text="First name", font=('Calibri', 16))
        self.firstnameEntry = tb.Entry(self.entryFrame, font=('Calibri', 16), validate="focus", validatecommand=(inputFunc, '%P'))
        self.firstnameLabel.grid(row=0, column=1, sticky="w", padx=(5, 5))
        self.firstnameEntry.grid(row=1, column=1, sticky="ew", padx=(5, 5), pady=(0, 5))
        # Surname
        self.surnameLabel = tb.Label(self.entryFrame, text="Surname", font=('Calibri', 16))
        self.surnameEntry = tb.Entry(self.entryFrame, font=('Calibri', 16), validate="focus", validatecommand=(inputFunc, '%P'))
        self.surnameLabel.grid(row=0, column=2, sticky="w", padx=(5, 5))
        self.surnameEntry.grid(row=1, column=2, sticky="ew", padx=(5, 5), pady=(0, 5))

        # Street
        self.streetLabel = tb.Label(self.entryFrame, text="Street", font=('Calibri', 16))
        self.streetEntry = tb.Entry(self.entryFrame, font=('Calibri', 16), validate="focus", validatecommand=(inputFunc, '%P'))
        self.streetLabel.grid(row=2, column=0, columnspan=3, sticky="w", padx=(5, 5))
        self.streetEntry.grid(row=3, column=0, columnspan=3, sticky="ew", padx=(5, 5), pady=(0, 5))

        # City
        self.cityLabel = tb.Label(self.entryFrame, text="City", font=('Calibri', 16))
        self.cityEntry = tb.Entry(self.entryFrame, font=('Calibri', 16), validate="focus", validatecommand=(inputFunc, '%P'))
        self.cityLabel.grid(row=4, column=0, sticky="w", padx=(5, 5))
        self.cityEntry.grid(row=5, column=0, sticky="ew", padx=(5, 5), pady=(0, 5))
        # State
        self.stateValue = StringVar()
        self.stateLabel = tb.Label(self.entryFrame, text="State", font=('Calibri', 16))
        self.stateDrop = tb.OptionMenu(self.entryFrame, self.stateValue, None, "ACT", "NSW", "NT", "QLD", "SA", "TAS", "VIC", "WA", bootstyle="light")
        self.stateLabel.grid(row=4, column=1, sticky="w", padx=(5, 5))
        self.stateDrop.grid(row=5, column=1, sticky="ew", padx=(5, 5), pady=(0, 5))
        # Postcode
        self.postcodeLabel = tb.Label(self.entryFrame, text="Postcode", font=('Calibri', 16))
        self.postcodeEntry = tb.Entry(self.entryFrame, font=('Calibri', 16), validate="focus", validatecommand=(inputFunc, '%P'))
        self.postcodeLabel.grid(row=4, column=2, sticky="w", padx=(5, 5))
        self.postcodeEntry.grid(row=5, column=2,sticky="ew", padx=(5, 5), pady=(0, 5))

        # Email
        self.emailLabel = tb.Label(self.entryFrame, text="Email", font=('Calibri', 16))
        self.emailEntry = tb.Entry(self.entryFrame, font=('Calibri', 16), validate="focus", validatecommand=(inputFunc, '%P'))
        self.emailLabel.grid(row=6, column=0, sticky="w", padx=(5, 5))
        self.emailEntry.grid(row=7, column=0, columnspan=3, sticky="ew", padx=(5, 5))
        
        # Position
        self.positionLabel = tb.Label(self.entryFrame, text="Job Position", font=('Calibri', 16))
        self.positionEntry = tb.Entry(self.entryFrame, font=('Calibri', 16), validate="focus", validatecommand=(inputFunc, '%P'))
        self.positionLabel.grid(row=8, column=0, columnspan=2, sticky="w", padx=(5, 5))
        self.positionEntry.grid(row=9, column=0, columnspan=2, sticky="ew", padx=(5, 5), pady=(0, 5))
        # Employment Type
        self.employValue = StringVar()
        self.employLabel = tb.Label(self.entryFrame, text="Employment Type", font=('Calibri', 16))
        self.employDrop = tb.OptionMenu(self.entryFrame,  self.employValue, None, "Casual", "Full-Time", "Part-Time", bootstyle="light")
        self.employLabel.grid(row=8, column=2, sticky="w", padx=(5, 5))
        self.employDrop.grid(row=9, column=2, sticky="ew", padx=(5, 5), pady=(0, 5))

        # Employer
        self.employerLabel = tb.Label(self.entryFrame, text="Employer (Payroll Services, etc.)", font=('Calibri', 16))
        self.employerEntry = tb.Entry(self.entryFrame, font=('Calibri', 16), validate="focus", validatecommand=(inputFunc, '%P'))
        self.employerLabel.grid(row=10, column=0, columnspan=3, sticky="w", padx=(5, 5))
        self.employerEntry.grid(row=11, column=0, columnspan=3, sticky="ew", padx=(5, 5), pady=(0, 5))

        # Manager
        self.managerLabel = tb.Label(self.entryFrame, text="Reports To", font=('Calibri', 16))
        self.managerEntry = tb.Entry(self.entryFrame, font=('Calibri', 16), validate="focus", validatecommand=(inputFunc, '%P'))
        self.managerLabel.grid(row=12, column=0, sticky="w", padx=(5, 5))
        self.managerEntry.grid(row=13, column=0, columnspan=3, sticky="ew", padx=(5, 5), pady=(0, 5))

        # Start Date
        self.startdateLabel = tb.Label(self.entryFrame, text="Start Date", font=('Calibri', 16))
        self.startEntry = tb.DateEntry(self.entryFrame, dateformat='%d/%m/%Y', bootstyle="light")
        self.startdateLabel.grid(row=14, column=0, sticky="w", padx=(5, 5))
        self.startEntry.grid(row=15, column=0, columnspan=3, sticky="w", padx=(5, 5), pady=(0, 5))

        # TEP
        self.tepLabel = tb.Label(self.entryFrame, text="TEP ($)", font=('Calibri', 16))
        self.tepEntry = tb.Entry(self.entryFrame, font=('Calibri', 16), validate="focus", validatecommand=(digitFunc, '%P'))
        self.tepLabel.grid(row=16, column=0, sticky="w", padx=(5, 5))
        self.tepEntry.grid(row=17, column=0, columnspan=2, sticky="ew", padx=(5, 5), pady=(0, 5))

        # Rate
        self.rateLabel = tb.Label(self.entryFrame, text="Super Rate (%)", font=('Calibri', 16))
        self.rateEntry = tb.Entry(self.entryFrame, font=('Calibri', 16), validate="focus", validatecommand=(digitFunc, '%P'))
        self.rateLabel.grid(row=16, column=2, sticky="w", padx=(5, 5))
        self.rateEntry.grid(row=17, column=2, sticky="ew", padx=(5, 5), pady=(0, 5))
        
        self.entryFrame.pack(padx=10, pady=10)

        self.buttonFrame = tk.Frame(self.root)

        # TODO: move buttons
        tb.Style().configure('TButton', font=('Calibri', 16))
        self.createButton = tb.Button(self.buttonFrame, text="Create", command=self.create, bootstyle="primary")
        self.createButton.grid(row=1, column=1, sticky="ew", padx=(5, 5), pady=(0, 5))

        self.resetButton = tb.Button(self.buttonFrame, text="Reset", command=self.reset, bootstyle="secondary")
        self.resetButton.grid(row=1, column=2, sticky="ew", padx=(5, 5), pady=(0, 5))

        self.buttonFrame.pack(padx=10, pady=10)

        self.root.protocol("WM_DELETE_WINDOW", self.onClose)
        self.root.mainloop()

    def onClose(self) -> None:
        if messagebox.askyesno(title="Quit?", message="Do you really want to quit?"):
            self.root.destroy()

    def create(self) -> None:
        entryValues = {
            "title": self.titleEntry.get(),
            "firstname": self.firstnameEntry.get(),
            "surname": self.surnameEntry.get(),
            "street": self.streetEntry.get(),
            "city": self.cityEntry.get(),
            "state": self.stateValue.get(),
            "postcode": self.postcodeEntry.get(),
            "email": self.emailEntry.get(),
            "position": self.positionEntry.get(),
            "employType": self.employValue.get(),
            "employer": self.employerEntry.get(),
            "manager": self.managerEntry.get(),
            "startDate": self.startEntry.entry.get(),
            "tep": self.tepEntry.get(),
            "rate": self.rateEntry.get(),
        }

        # validation of entries
        for valueType in entryValues:
            if not self.validateInput(entryValues[valueType]):
                messagebox.showinfo("Error", self.getString(valueType) + " is missing input!")
                return
            
            if (valueType == "tep" or valueType == "rate") and not self.validateNumber(entryValues[valueType]):
                messagebox.showinfo("Error", self.getString(valueType) + " must be a number!")
                return

            entryValues[valueType] = entryValues[valueType].strip()

        templateFilepath = self.resource_path('template.docx')
        if not os.path.exists(templateFilepath):
            messagebox.showerror("Error", "Contract template file is missing!")
            return
            
        template = Document(templateFilepath)

        # takes entryValues as String
        template = CreateContract(template, entryValues)

        messagebox.showinfo("Message", "Contract successfully created!")

    def reset(self) -> None:
        if not messagebox.askyesno(title="Reset?", message="Do you really want to reset the entries?"):
            return
            
        self.titleEntry.delete(0, tk.END)
        self.firstnameEntry.delete(0, tk.END)
        self.surnameEntry.delete(0, tk.END)
        self.streetEntry.delete(0, tk.END)
        self.cityEntry.delete(0, tk.END)
        self.stateValue.set("")
        self.postcodeEntry.delete(0, tk.END)
        self.emailEntry.delete(0, tk.END)
        self.positionEntry.delete(0, tk.END)
        self.employValue.set("")
        self.employerEntry.delete(0, tk.END)
        self.managerEntry.delete(0, tk.END)
        
        self.startEntry = tb.DateEntry(self.entryFrame, dateformat='%d/%m/%Y', bootstyle="light")
        self.startEntry.grid(row=15, column=0, columnspan=3, sticky="w", padx=(5, 5), pady=(0, 5))
        
        self.tepEntry.delete(0, tk.END)
        self.rateEntry.delete(0, tk.END)   

    def validateNumber(self, x) -> bool:
        try:
            float(x)
            return True
        except ValueError:
            return False
    
    def validateInput(self, s) -> bool:
        return s.strip() != ''

    def getString(self, s) -> str:
        match s:
            case "title":
                return "Title"
            case "firstname":
                return "First name"
            case "surname":
                return "Surnanme"
            case "street":
                return "Street"
            case "city":
                return "City"
            case "state":
                return "State"
            case "postcode":
                return "Postcode"
            case "email":
                return "Email"
            case "position":
                return "Job position"
            case "employer":
                return "Employer"
            case "startDate":
                return "Start date"
            case "employType":
                return "Employment type"
            case "tep":
                return "TEP"
            case "rate":
                return "Rate"
            case "manager":
                return "\'Reports to\'"
    
    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

MyGUI()