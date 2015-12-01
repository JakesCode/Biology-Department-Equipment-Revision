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
		l = Label(master, text="Biology Department Equipment Revision", font=("Courier New", 22)).grid(row=0,column=1, columnspan=1, rowspan=1)

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


		self.dateEntry = Entry(master, font=("Courier New", 11), width=40)
		self.dateEntry.grid(row=1,column=1)
		self.dateEntry.delete(0, END)
		self.dateEntry.insert(0, "Enter Date Here....")

		self.dateVar = StringVar(master)
		self.dateVar.set("Please remember to add a \n0 if the date is single-digit at any point.")
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


		self.equipment = Text(master, width=50, font=("Courier New", 13))
		self.equipment.grid(row=3,column=1)


		self.hazcardsLabel = Label(master, text="Hazcards referred to....", font = ("Courier New", 13)).grid(row=4,column=0)

		self.hazcards = Entry(master, width=75, font=("Courier New", 13))
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
			box = messagebox.askquestion("Send", ("Are you sure you want to send the email?\nMake sure that you are " + self.teacherVar.get() + " and you need your requested items on " + self.dateEntry.get() + " during " + self.periodVar.get() + "."), icon="question")
			# All good, let's go! #
			if box == 'yes':
				self.sendEmail()

	def sendEmail(self):
		recipient = "technicians@gsal.org.uk"
		subject = ("Practical Request, " + str(self.dateEntry.get()) + " from " + self.teacherVar.get())
		self.equipmentEmail = self.equipment.get("1.0", "end-1c").replace("\n", "<br>")

		self.bodyText = ("<h2>Practical Request from <u>" + "<b>" + self.teacherVar.get() + "</b>" + "</u>.</h2><br><br><h3>" + self.teacherVar.get() + " has requested the following equipment:</b3><br>" + "<h1 style='color:blue'>" + self.equipmentEmail + "</h1><br>" + "<h1>It is needed during " + self.periodVar.get() + " on <u>" + self.dateEntry.get() + "</u>.</h1>")

		outlook = win32.Dispatch('outlook.application')
		mail = outlook.CreateItem(0)
		mail.To = recipient
		mail.Subject = subject
		mail.HtmlBody = self.bodyText
		mail.Display(True)

		self.setReminder()

		messagebox.showinfo("Success!", "Your request/reminder has been sent successfully.")


	def setReminder(self):

		self.finalisedDatePreFormatted = (self.dateEntry.get().replace("/", "-"))

		# Americanise the date, because Outlook #
		self.finalisedDate = (self.finalisedDatePreFormatted[3:] + "-" + self.finalisedDatePreFormatted[:2])

		self.finalisedPeriod = ""

		if self.periodVar.get() == "Period 1":
			self.finalisedPeriod = " 9:00"
		elif self.periodVar.get() == "Period 2":
			self.finalisedPeriod = " 9:55"
		elif self.periodVar.get() == "Period 3":
			self.finalisedPeriod = " 11:05"
		elif self.periodVar.get() == "Period 4":
			self.finalisedPeriod = " 12:00"
		elif self.periodVar.get() == "Period 5":
			self.finalisedPeriod = " 14:10"
		elif self.periodVar.get() == "Period 6":
			self.finalisedPeriod = " 15:05"

		start = '2015-' + self.finalisedDate + self.finalisedPeriod
		subject = 'Practical Request from ' + self.teacherVar.get()

		# Set the actual reminder #
		self.addevent(start, subject)


	def addevent(self, start, subject):
		oOutlook = win32.Dispatch("Outlook.Application")
		appointment = oOutlook.CreateItem(1)
		appointment.Start = start
		appointment.Subject = subject
		appointment.Duration = 50
		appointment.Location = (self.finalisedPeriod + "/" + self.periodVar.get())
		appointment.ReminderSet = True
		appointment.ReminderMinutesBeforeStart = 10
		appointment.Save()
 

root = Tk()
app = AppWindow(root)
root.iconbitmap(r"icons/favicon.ico")
root.mainloop()
