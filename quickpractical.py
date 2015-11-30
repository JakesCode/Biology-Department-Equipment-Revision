# QuikPractical! System #
# By Jake Stringer 2015 #

import smtplib
import datetime
import win32com.client as win32
import os
from tkinter import *
from tkinter import messagebox

class AppWindow():
	def __init__(self, master):
		self.master = master
		self.master.title("Biology Department Equipment Revision")
		self.build(master)

	def build(self, master):
		l = Label(master, text="Biology Department Equipment Revision", font=("Courier New", 22, "bold")).grid(row=0,column=1, columnspan=1, rowspan=1)

		self.teachersDict = {"Dr. Gilbert [ALG]":"ALG",
			"Dr. Pett [MRP]":"MRP",
			"Mr. Smith [MS]":"MS",
			"Mrs. Allison [CGA]":"CGA",
			"Mrs. Baxter [AMB]":"AMB",
			"Mrs. Brown [LVB]":"LVB",
			"Mrs. Jagger [CHJ]":"CHJ",
			"Mrs. Walker [BAW]":"BAW",
			"Mrs. Yuasa [AY]":"AY"}
		self.teacherVar = StringVar(master)
		self.teachersList = list(self.teachersDict.keys())
		self.teachersList.sort()
		self.teachersList.insert(0,"Which Teacher?")
		self.teacherVar.set(self.teachersList[0])

		self.teacherSelect = OptionMenu(master, self.teacherVar, *self.teachersList)
		self.teacherSelect.grid(row=1,column=0, columnspan=1, rowspan=1)


		self.dateEntry = Entry(master)
		self.dateEntry.grid(row=1,column=1)
		self.dateEntry.delete(0, END)
		self.dateEntry.insert(0, "Enter Date Here....")

		self.dateVar = StringVar(master)
		self.dateVar.set("Today's date is " + str(datetime.date.today()))
		self.dateSuffix = Label(master, text=self.dateVar.get()).grid(row=1,column=2)


		self.numbersList = ["Number of groups?"]
		for x in range(0, 10):
			self.numbersList.append("I want " + str(x+1) + " groups.")

		self.groupsVar = StringVar(master)
		self.groupsVar.set(self.numbersList[0])

		self.groupsSelect = OptionMenu(master, self.groupsVar, *self.numbersList)
		self.groupsSelect.grid(row=2,column=0)


		self.periodList = ["Which period?"]
		for y in range(0, 6):
			self.periodList.append("Period " + str(y+1))

		self.periodVar = StringVar(master)
		self.periodVar.set(self.periodList[0])

		self.periodSelect = OptionMenu(master, self.periodVar, *self.periodList)
		self.periodSelect.grid(row=2,column=1)


		self.yearList = ["Which year group?"]
		for z in range(7, 14):
			self.yearList.append("Year " + str(z))

		self.yearVar = StringVar(master)
		self.yearVar.set(self.yearList[0])

		self.yearSelect = OptionMenu(master, self.yearVar, *self.yearList)
		self.yearSelect.grid(row=2,column=2)


		self.equipmentLabel = Label(master, text="Equipment Needed....", font = ("Courier New", 15)).grid(row=3,column=0,columnspan=1,rowspan=1)


		self.equipment = Text(master, width=50, font=("Courier New", 13), image=None)
		self.equipment.grid(row=3,column=1)


		self.hazcardsLabel = Label(master, text="Hazcards referred to....", font = ("Courier New", 13)).grid(row=4,column=0)

		self.hazcards = Entry(master, width=75)
		self.hazcards.grid(row=4,column=1)


		self.isChecked = IntVar()
		self.riskAssessment = Checkbutton(master, text="I have carried out a risk assessment.", variable=self.isChecked, font=("Courier New", 15))
		self.riskAssessment.grid(row=5,column=1)


		self.sendButton = Button(master, text="Submit", font=("Courier New", 15), bg="#B8B0B0", command=lambda:self.validateSubmit(), width=40)
		self.sendButton.grid(row=6,column=1)


	def validateSubmit(self):
		listOfErrors = []

		if self.teacherVar.get() == "Which Teacher?":
			listOfErrors.append("Please select a teacher.\n")

		if self.dateVar.get() == "":
			listOfErrors.append("Please enter a date.\n")

		if self.groupsVar.get() == "Number of groups?":
			listOfErrors.append("Please enter how many groups you want.\n")

		if self.periodVar.get() == "Which period?":
			listOfErrors.append("Please enter a period.\n")

		if self.yearVar.get() == "Which year group?":
			listOfErrors.append("Please enter a year group.\n")

		if self.equipment.get("1.0","end-1c") == "":
			listOfErrors.append("You haven't requested any equipment.\n")

		if self.hazcards.get() == "":
			listOfErrors.append("You haven't specified any hazcards.\n")

		if self.isChecked.get() == 0:
			listOfErrors.append("You haven't agreed to have done a risk assessment.")

		listOfErrorsString = ""
		for x in listOfErrors:
			listOfErrorsString += x

		if len(listOfErrors) > 0:
			messagebox.showwarning("Error - Can't Proceed", ("*********"*8 + "\nThe email couldn't send. The following issues have been detected: \n" + "*********"*8 + "\n" + listOfErrorsString))
		else:
			messagebox.askquestion("Send", ("Are you sure you want to send the email?\nMake sure that you are " + self.teacherVar.get() + " and you need your requested items on " + self.dateEntry.get() + " during " + self.periodVar.get() + "."), icon="question")
			# All good, let's go! #
			self.sendEmail()

	def sendEmail(self):
		recipient = "technicians@gsal.org.uk"
		subject = ("Practical Request, " + str(self.dateEntry.get()) + " from " + self.teacherVar.get())
		text = ("<h2><u>Practical Request from " + "<b>" + self.teacherVar.get() + "</b>" + ".</u></h2><br><br><h2>" + self.teacherVar.get() + " has requested the following:</h2><br>" + "<h1>" + self.equipment.get("1.0","end-1c") + "</h1>")

		outlook = win32.Dispatch('outlook.application')
		mail = outlook.CreateItem(0)
		mail.To = recipient
		mail.Subject = subject
		mail.HtmlBody = text
		mail.Display(True)

		self.createEvent(2015, (self.equipment.get("1.0","end-1c")))

		messagebox.showinfo("Success!", "Your request has been sent successfully.")


	def createEvent(self, start, subject):
		import win32com.client
		oOutlook = win32com.client.Dispatch("Outlook.Application")
		appointment = oOutlook.CreateItem(1) # 1=outlook appointment item
		appointment.Start = start
		appointment.Subject = subject
		appointment.Duration = 20
		appointment.Location = "GSAL"
		appointment.ReminderSet = True
		appointment.ReminderMinutesBeforeStart = 0
		appointment.Save()
		return

		table = {"11-16":20, "12-1":30, "12-16":40, "12-31":50}

		for item in table.keys():
			start = '2015-' + item + ' 18:35'
			subject = 'P-bars. To do:' + str(table[item])
			addevent(start, subject)


root = Tk()
app = AppWindow(root)
root.mainloop()
