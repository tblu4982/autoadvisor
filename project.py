from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.edge.options import Options
import tkinter #GUI
from tkinter import *
from tkinter import messagebox
from datetime import datetime
import os #used for read/write to files
import re #used for regular expressions
import sys #used to stop execution under certain circumstances

#Class to hold user login info for session
class credentials():
    #Opens a window and prompts user to select if they are student or advisor
    def greeting_window (self, greeting):
        portal = Tk()
        portal.geometry("200x70")
        portal.title("Auto Advisor")
        welcome = Label(portal, text = greeting)
        prompt = Label(portal, text='Are you an Advisor or a Student?')
        advisor = Button(portal, text="Advisor", command = lambda:[portal.destroy(), self.click_advisor()])
        student = Button(portal, text="Student", command = lambda:[portal.destroy(), self.click_student()])

        welcome.pack(side = TOP)
        prompt.pack(side = TOP)
        advisor.pack(side = LEFT, expand = True, fill = BOTH)
        student.pack(side = LEFT, expand = True, fill = BOTH)
        portal.mainloop()

    #method to grab student login info
    def click_student(self):
        print("Clicked 'Student'")
        self.auth = "student"
        #opens a window to grab student login info
        login = Tk()
        login.geometry("200x90")
        login.title("Auto Advisor - Student")
        welcome = Label(login, text = "Loggin in as: Student")
        Label(login, text='User ID').grid(row=1)
        Label(login, text='Password').grid(row=2)
        uid = Entry(login)
        pwd = Entry(login, show ="*")
        #when submit button is clicked, it sends credentials to Banner Portal
        submit = Button(login, text='Submit', command = lambda:[self.set_credentials(uid, pwd), login.destroy()])
        back = Button(login, text='Return', command = lambda:[login.destroy(), self.greeting_window(self.greeting)])
        welcome.grid(row = 0, columnspan = 3)
        uid.grid(row = 1, column = 1, columnspan = 2)
        pwd.grid(row = 2, column = 1, columnspan = 2)
        submit.grid(row = 3, column = 1)
        back.grid(row = 3, column = 2)
        
    #sets user's login info
    def set_credentials(self, uid, pwd):
        self.username = uid.get()
        self.password = pwd.get()
        
    #returns login info
    def get_credentials(self):
        return self.auth, self.username, self.password
    
    #method to grab advisor login info
    def click_advisor(self):
        print("Clicked 'Advisor'")
        self.auth = "advisor"
        #opens a window to grab student login info
        login = Tk()
        login.geometry("200x90")
        login.title("Auto Advisor - Advisor")
        welcome = Label(login, text = "Logging in as: Advisor")
        Label(login, text='User ID').grid(row=1)
        Label(login, text='Password').grid(row=2)
        uid = Entry(login)
        pwd = Entry(login, show ="*")
        #when submit button is clicked, it sends credentials to Banner Portal
        submit = Button(login, text='Submit', command = lambda:[self.set_credentials(uid, pwd), login.destroy()])
        back = Button(login, text='Return', command = lambda:[login.destroy(), self.greeting_window(self.greeting)])
        welcome.grid(row = 0, columnspan = 3)
        uid.grid(row = 1, column = 1, columnspan = 2)
        pwd.grid(row = 2, column = 1, columnspan = 2)
        submit.grid(row = 3, column = 1)
        back.grid(row = 3, column = 2)

    #grabs student id if advisor is logged in
    def get_student_id(self):
        get_sid = Tk()
        get_sid.geometry("230x80")
        get_sid.title("Auto Advisor")
        prompt = Label(get_sid, text = "Enter Student V-Number")
        Label(get_sid, text='V-Number:').grid(row = 1, column = 0)
        sid = Entry(get_sid)
        submit = Button(get_sid, text='Submit', command = lambda:[self.set_sid(sid), get_sid.destroy()])
        cancel = Button(get_sid, text='Cancel', command = get_sid.destroy)
        prompt.grid(row = 0, column = 1, columnspan = 2)
        sid.grid(row = 1, column = 1, columnspan = 2)
        submit.grid(row = 2, column = 1)
        cancel.grid(row = 2, column = 2)
        get_sid.mainloop()
        
    #returns student id
    def set_sid(self, vnum):
        self.sid = vnum.get()
        
    def return_sid(self):
        return self.sid

    #method that prompts user to correct invalid login info
    def invalid_creds(self):
        self.greeting = 'Invalid login credentials!'
        self.greeting_window(self.greeting)

    #instantiates class
    def __init__(self):
        self.username = ""
        self.password = ""
        self.auth = ""
        self.sid = ""
        self.greeting = 'Welcome to Auto-Advisor!'
        self.greeting_window(self.greeting)

#method to return appropriate value for term dropdown
def get_timecode():
    #get current month and year
    month = datetime.now().month
    year = str(datetime.now().year)
    #change month to string
    if month >= 1:
        #set month to "01" if month is between 1 and 4
        if month >= 5:
            #set month to "05" if month is between 5 and 7
            if month >= 8:
                #set month to "08" if month is between 8 and 11
                if month == 12:
                    #change month to string if month is 12
                    month = "12"
                else:
                    month = "08"
            else:
                month = "05"
        else:
            month = "01"
    #return year and month to use as value for dropdown selector
    return year + month
        
#method to create file to store data
def create_file_path(fullname ,path, filename, file_type, array):
    #check if file directory exists
    if os.path.exists(path):
        #if folder exists, check if file already exists
        file_path = path + filename
        #if file exists, overwrite it
        if os.path.exists(file_path):
            #remove old file, then create new file
            print("Overwriting " + file_type + " file for " + fullname + "...")
            os.remove(file_path)
            with open(file_path, 'w') as target_file:
                if file_type == "courses":
                    for line in array:
                        for index in line:
                            target_file.write(index + " ")
                        target_file.write("\n")
                else:
                    for line in array:
                        target_file.write(line + "\n")
        else:
            #create new file
            print("Creating " + file_type + " file for " + fullname + "...")
            with open(file_path, 'w') as target_file:
                if file_type == "courses":
                    for line in array:
                        for index in line:
                            target_file.write(index + " ")
                        target_file.write("\n")
                else:
                    for line in array:
                        target_file.write(line + "\n")
    else:
        print("Error! Folder path for " + fullname + " does not exist!")

#Gets user login credentials
user = credentials()
try:
    auth, username, password = user.get_credentials()
#Exception Handling that closes program if tkinter box is closed prematurely
except AttributeError:
    sys.exit("Program Terminated")

if len(auth) == 0 or len(username) == 0 or len(password) == 0:
    sys.exit("Credentials not set")

#--------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------
#May be relevant: SessionNotCreatedException
options = Options()
#--------------COMMENT OUT WHEN TESTING OR DEBUGGING---------------
options.add_argument("headless")
#--------------COMMENT OUT WHEN TESTING OR DEBUGGING---------------
driver = webdriver.Edge(options = options)
driver.get('https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_WWWLogin')
#--------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------

#enter credentials into Banner
uid = driver.find_element(By.ID, "UserID")
pwd = driver.find_element(By.ID, "PIN")

uid.send_keys(username)
pwd.send_keys(password)

driver.find_element(By.XPATH, "//form").submit()

#--------ADD LOGIC THAT CHECKS TO SEE IF USER IS AT LANDING PAGE------
#LANDING PAGE URL: https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_GenMenu?name=bmenu.P_MainMnu
#use regular expressions since end of url can vary between users
home_url = "https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_GenMenu?name=bmenu.P_MainMnu"
url = driver.current_url.split("&")[0]
print(url)
#try:
    #auth, username, password = user.invalid_creds()
#Exception Handling that closes program if tkinter box is closed prematurely
#except AttributeError:
    #sys.exit("Program Terminated")
#--------ADD LOGIC THAT CHECKS TO SEE IF USER IS AT LANDING PAGE------

fullname = ""

if auth == "advisor":
    #crawl through banner until we get to student id
    driver.find_element(By.LINK_TEXT, "Faculty and Advisors").click()
    driver.find_element(By.LINK_TEXT, "Student Information Menu").click()
    driver.find_element(By.LINK_TEXT, "ID Selection").click()
    #The following dropdown menu lacks an id
    #so to select for current semester, I noticed html code stored
    #dropdown values as dates, I used a method to grab current
    #month and year, and store it to a variable to use as value
    date = get_timecode()
    term = Select(driver.find_element(By.XPATH, "//select[@name='term']"))
    term.select_by_value(date)
    driver.find_element(By.XPATH, "//td[@class='dedefault']").submit()
    #requests student id and uses it to get to transcript
    user.get_student_id()
    sid = user.return_sid()
    vnum = driver.find_element(By.ID, "Stu_ID")
    vnum.send_keys(sid)
    driver.find_element(By.XPATH, "//table[contains(@class, 'dataentrytable')]").submit()
    #gets student name
    name = driver.find_element(By.TAG_NAME, 'b')
    fullname = name.text.replace(" ", "").replace(".", "")
    driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()

    #crawl webpage to academic transcript
    driver.find_element(By.LINK_TEXT, "Academic Transcript").click()

    levl_id = Select(driver.find_element(By.ID, "levl_id"))
    type_id = Select(driver.find_element(By.ID, "tprt_id"))

if auth == "student":
    #grab person's name to use for file name
    name = driver.find_elements(By.XPATH, "//td[@class='pldefault']")
    for i in name:
        if i.text[0:7] == "Welcome":
            fullname = i.text.split(",")[1].replace(" ", "").replace(".", "")

    #crawl webpage to academic transcript
    driver.find_element(By.LINK_TEXT, "Student").click()
    driver.find_element(By.LINK_TEXT, "Student Records").click()

    #crawl webpage to academic transcript
    driver.find_element(By.LINK_TEXT, "Academic Transcript").click()

    levl_id = Select(driver.find_element(By.ID, "levl_id"))
    type_id = Select(driver.find_element(By.ID, "type_id"))

#create folder using student name
#check if name has been established
if not len(fullname) == 0:
    #search to see if student folder already exists
    path = "students/" + fullname
    if not os.path.exists(path):
        os.makedirs(path)
        if os.path.exists(path):
            print("Path successfully created!")
        else:
            print("Failed to create path!")
    else:
        print("Path already exists!")
else:
    sys.exit("An unexpected error has occurred: Unable to locate student's name!")

driver.find_element(By.XPATH, "//table[contains(@class, 'dataentrytable')]").submit()

#scrape transcript for courses and semesters
output = driver.find_elements(By.XPATH, "//tr")

courses = []
semesters = []
#boolean flags to distinguish semesters
sem_start = sem_end = False
#used to determine if we reached current/future semesters
is_curr = False

for i in output:
    #Find each semester as we iterate through scraped data
    #signifies the start of courses in progress
    if re.search("COURSES IN PROGRESS", i.text) and not is_curr:
        is_curr = True
    if re.search("Spring", i.text) or re.search("Summer", i.text) or \
       re.search("Fall", i.text) or re.search("Winter", i.text):
        if sem_start:
            #lets us know that we've reached the end of a previous semester
            sem_end = True
        #signifies the start of a new semester
        sem_start = True
        #store the current semester into semesters array
        semesters.append(i.text)
        semesters.append("-")
    #if we are in a semester, find the courses
    elif sem_start:
        #add marker to course array to signify semesters
        if sem_end:
            courses.append("-")
            sem_end = False
        #find a course and store it to the holder array
        if re.search("[A-Z][A-Z][A-Z][A-Z] ", i.text) and not re.search("-Top-", i.text):
            proto = i.text.replace("\n", " ").split(" ")
            #check if we reached current/future semesters
            if is_curr:
                #pop excess list indices
                while not re.search("[0-9]", proto[-1]):
                    proto.pop()
                #add 'inprog' to courses in progress
                proto.insert(-1, "inprog")
            #we reach this if we are not in current/future semesters
            else:
                while not re.search("[0-9]", proto[-1]):
                    proto.pop()
                #past courses have one more filled index than current/future sems
                proto.pop()
            #append the holder array to courses
            courses.append(proto)

#Web driver is no longer needed
driver.quit()
#Append separator for final course structure
courses.append("-")
    
#print courses to file
create_file_path(fullname ,path, "/courses.txt", "courses", courses)
#print semesters to file
create_file_path(fullname , path, "/semesters.txt", "semesters", semesters)
