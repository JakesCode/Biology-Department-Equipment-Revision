# QuikPractical! System #
# By Jake Stringer 2015 #

import smtplib
import win32com.client
import datetime
from tkinter import *
from tkinter import messagebox

class AppWindow():
	def __init__(self, master):
		self.master = master
		self.master.title("Biology Department Equipment Revision")
		self.build(master)

	def build(self, master):
		l = Label(master, text="Biology Department Equipment Revision", font=("Courier New", 22)).grid(row=0,column=1, columnspan=1, rowspan=1)

		self.teachersDict = {"Dr. Gilbert [ALG]":"alg@gsal.org.uk",
			"Dr. Pett [MRP]":"mrp@gsal.org.uk",
			"Mr. Smith [MS]":"ms@gsal.org.uk",
			"Mrs. Allison [CGA]":"cga@gsal.org.uk",
			"Mrs. Baxter [AMB]":"amb@gsal.org.uk",
			"Mrs. Brown [LVB]":"lvb@gsal.org.uk",
			"Mrs. Jagger [CHJ]":"chj@gsal.org.uk",
			"Mrs. Walker [BAW]":"baw@gsal.org.uk",
			"Mrs. Yuasa [AY]":"ay@gsal.org.uk"}
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


		self.equipment = Text(master, width=50, font=("Courier New", 13))
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
			messagebox.showwarning("Error - Can't Proceed", listOfErrorsString)
		else:
			messagebox.askquestion("Send", "Are you sure you've finished?", icon="question")
			# All good, let's go! #
			self.sendEmail()

	def sendEmail(self):
		# fromaddr = (self.teachersDict[self.teacherVar.get()])
		fromaddr = "test@gsal.org.uk"
		toaddrs = "technicians@gsal.org.uk"
		msg = "\r\n".join([
			"From: user",
			"To: user",
			"Subject: Practical Request from banternaut",
			"",
			"Lots of banter over here"
			])

		server = smtplib.SMTP_SSL("exchange.gsal.org")
		server.ehlo()
		# server.login(username,password)
		server.sendmail(fromaddr, toaddrs, msg)
		server.quit()
		messagebox.showinfo("Success!", "Your request has been sent successfully.")


root = Tk()
app = AppWindow(root)
root.mainloop()