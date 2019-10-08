
import getpass
global sheet, register 
import time 

class Credentials:
	global user, passwrd
	user = "user"
	passwrd = "User" 
	
	def logIn(self,n=2):
		print "LogIn Please \n"
		username = raw_input("UserName: ")
		password = getpass.getpass( )
		if (n==0):
			self.attempt()
		elif(username != user or password != passwrd):
			print "\n Invalid credentials \n"
			self.logIn(n-1)
		else:
			print "\n Logged in successfully! \n\n" 
		
	def attempt(self):	
		import ctypes
		print "Log In attempt failed \n Signing out",
		for x in range(3):
			time.sleep(0.5)
			print ".",
		ctypes.windll.user32.LockWorkStation()
		quit()
		
class Attendance:
	def attendanceMark(self,sheet):
		n= (sheet.max_row)
		row_n=0
		i=1
		for i in range (2,n+1):
			print "\n Roll No.:", ((sheet.cell(row=i, column=1)).value),((sheet.cell(row=i, column=2)).value)
			marker(sheet,i)	
			attended=(sheet.cell(row=i, column=3).value)
			total=(sheet.cell(row=i, column=4).value)
			sheet.cell(row=i, column=5).value=attendancePercentage(attended,total)
		register.save(file)
		raw_input("Press return to Main Menu")
		choice()
	def update(self,sheet):
		n= (sheet.max_row)
		row_n=0
		i=1
		for i in range (2,n+1):	
			attended=(sheet.cell(row=i, column=3).value)
			total=(sheet.cell(row=i, column=4).value)
			sheet.cell(row=i, column=5).value=attendancePercentage(attended,total)
		register.save(file)
	def attendancePercentage(a,t):
		m=(float(a)/t)*100
		return round(m,2)
	def marker(sheet,i):
		att=raw_input("Mark the attendance(p/a): ")
		if (att=="p" or att=="P"):
			(sheet.cell(row=i, column=3).value) +=1
			(sheet.cell(row=i, column=4).value) +=1
		elif (att=="a" or att=="A"):
			(sheet.cell(row=i, column=4).value) +=1
		else:
			print "Invalid! Mark Correctly"
			marker(sheet,i)

def studentsList(sheet):
	print "\n  Roll No.			Name			Attendance(%)"
	for i in range (2,(sheet.max_row+1)):
		print "  ",((sheet.cell(row=i, column=1)).value),"				",((sheet.cell(row=i, column=2)).value),"			",((sheet.cell(row=i, column=5)).value)
	raw_input("Press return to quit")
	choice()

def choice():
	print "\n Please choose the following options: \n1.Mark the Attendance \n2.List of Students \n3.Change Username & Password \n4.Exit"	
	c= input("\n Enter your choice:")
	if(c!=1 and c!=2 and c!=3 and c!=4): 
		print "Invalid choice"
		choice()
	else:
		if (c==1):
			student.attendanceMark(sheet)
		elif (c==2):
			student.update(self,sheet)
			studentsList(sheet)
		elif (c==4):
			print "Saving",
			for i in range (3):
				time.sleep(0.5)
				print ".",
			print "\nClosing!"
			time.sleep(0.5)
			quit()
			
def fileOpen():
	file=raw_input("Enter the path of file:")
	try:
		register = openpyxl.load_workbook(file)
		sheet = register.active
		return sheet
	except:
		print "\n Invalid File Format!! \n"
		fileOpen()

def consoleSize():
	from ctypes import windll, create_string_buffer

	h = windll.kernel32.GetStdHandle(-12)
	csbi = create_string_buffer(22)
	res = windll.kernel32.GetConsoleScreenBufferInfo(h, csbi)

	if res:
		import struct
		(bufx, bufy, curx, cury, wattr,
		left, top, right, bottom, maxx, maxy) = struct.unpack("hhhhHhhhhhh", csbi.raw)
		sizex = right - left + 1
		sizey = bottom - top + 1
	else:
		sizex, sizey = 80, 25 
	
	return sizex

student = Attendance()

login_id = Credentials()

char = "Welcome to your Attendance Assistant"
print (char.center(consoleSize()))

print ((time.ctime()).center(consoleSize()+50))

login_id.logIn()

sheet=fileOpen()

choice()	

