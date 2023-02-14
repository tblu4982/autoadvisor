from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import tkinter #GUI
from tkinter import *
from tkinter import messagebox
from datetime import datetime
import os #used for read/write to files
import re #used for regular expressions
import sys #used to stop execution under certain circumstances
import preprocess

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
        welcome = Label(login, text = "Logging in as: Student")
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

    def login_warning(self):
        warn = Tk()
        warn.geometry("200x90")
        warn.title("Warning!")
        warning = Label("Warning! You have failed to login 3 times now. " + \
                        "Please ensure that you enter your information correctly.")

    def login_timeout(self):
        pass

    #method that prompts user to correct invalid login info
    def invalid_creds(self):
        self.greeting = 'Invalid login credentials!'
        self.login_count += 1
        if self.login_count < 3:
            self.greeting_window(self.greeting)
        elif self.login_count == 3:
            self.login_warning()
        else:
            self.login_timeout()

    #instantiates class
    def __init__(self):
        self.username = ""
        self.password = ""
        self.auth = ""
        self.sid = ""
        self.greeting = 'Welcome to Auto-Advisor!'
        self.greeting_window(self.greeting)
        self.login_count = 0

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

def build_path(path):
    #create folder using student name
    #check if name has been established
    if bool(fullname):
        #search to see if student folder already exists
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

def build_files(path, driver, fullname):
    driver.find_element(By.XPATH, "//table[contains(@class, 'dataentrytable')]").submit()

    #scrape transcript for courses and semesters
    data = driver.find_elements(By.XPATH, "//table[@class = 'datadisplaytable']")

    courses = []
    semesters = []
    #boolean flags to distinguish semesters
    sem_start = sem_end = False
    #used to determine if we reached current/future semesters
    is_curr = False
    #used to hold course as we build it from data
    proto = ""

    for i in data:
        output = i.text.split('\n')

    for line in output:
        #Find each semester as we iterate through scraped data
        #signifies the start of courses in progress
        if re.search("COURSES IN PROGRESS", line) and not is_curr:
            is_curr = True
        if re.search("Spring", line) or re.search("Summer", line) or \
           re.search("Fall", line) or re.search("Winter", line) or \
           re.search("1:", line):
            if sem_start:
                #lets us know that we've reached the end of a previous semester
                sem_end = True
            #signifies the start of a new semester
            sem_start = True
            #store the current semester into semesters array
            if re.search("Spring [0-9]{4}|Summer [0-9]{4}|Fall [0-9]{4}|Winter [0-9]{4}", line):
                semesters.append(re.search("Spring [0-9]{4}|Summer [0-9]{4}|Fall [0-9]{4}|Winter [0-9]{4}", line).group())
            #for semesters that are missing a term name
            else:
                semesters.append("UNKNOWN")
            semesters.append("-")
        #if we are in a semester, find the courses
        elif sem_start:
            #add marker to course array to signify semesters
            if sem_end:
                courses.append("-")
                sem_end = False
            #append credits to course
            if bool(proto):
                proto.append(line)
                course = proto.copy()
                #pop empty indices at beginning of course
                while not bool(course[0]):
                    course.pop(0)
                #remove 'U' from courses
                if course[2] == 'U':
                    course.pop(2)
                #append course to array and reset proto
                courses.append(course)
                proto.clear()
            #find a course and store it to the holder array
            elif re.search("[A-Z]{4} ", line) and not re.search("-Top-", line):
                proto = line.replace("\n", " ").split(" ")
                #check if we reached current/future semesters
                if is_curr:
                    #add 'inprog' to current/future courses
                    proto.append("inprog")

    #Append separator for final course structure
    courses.append("-")

    #print courses to file
    create_file_path(fullname , path, "/courses.txt", "courses", courses)
    #print semesters to file
    create_file_path(fullname , path, "/semesters.txt", "semesters", semesters)
        
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

if not bool(auth) or not bool(username) or not bool(password):
    sys.exit("Credentials not set")

#--------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------
#May be relevant: SessionNotCreatedException
driver = webdriver.Edge()
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
#except AttributeError:
    #sys.exit("Program Terminated")
#--------ADD LOGIC THAT CHECKS TO SEE IF USER IS AT LANDING PAGE------

#Used to hold student V-Numbers
vnums = []
try:
    f1 = open("vnums.txt")
    for vnum in f1:
        vnums.append(vnum)
    f1.close()
except FileNotFoundError:
    pass
#Used to hold student names
fullname = []

itr = 0
if auth == "advisor":
    #crawl through banner until we get to student id
    driver.find_element(By.LINK_TEXT, "Faculty and Advisors").click()
    for vnum in vnums:
        driver.find_element(By.LINK_TEXT, "Student Information Menu").click()
        driver.find_element(By.LINK_TEXT, "ID Selection").click()
        #The following dropdown menu lacks an id
        #so to select for current semester, I noticed html code stored
        #dropdown values as dates, I used a method to grab current
        #month and year, and store it to a variable to use as value
        date = get_timecode()
        #We visit the 'Select Term' page only once, during the first time we crawl
        if vnum == vnums[0]:
            term = Select(driver.find_element(By.XPATH, "//select[@name='term']"))
            term.select_by_value(date)
            driver.find_element(By.XPATH, "//td[@class='dedefault']").submit()
        #requests student id and uses it to get to transcript
        sid = driver.find_element(By.ID, "Stu_ID")
        sid.send_keys(vnum)
        #Additional crawling logic if we've reached the last vnum in list
        if vnum == vnums[-1]:
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
            #driver.find_element(By.XPATH, "//table[contains(@class, 'dataentrytable')]").submit()
            #gets student name
            name = driver.find_element(By.TAG_NAME, 'b')
            fullname.append(name.text.replace(" ", "").replace(".", ""))
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
        #For all other vnums that are not last in list
        else:
            name = driver.find_element(By.TAG_NAME, 'b')
            fullname.append(name.text.replace(" ", "").replace(".", ""))
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()

        #crawl webpage to academic transcript
        driver.find_element(By.LINK_TEXT, "Academic Transcript").click()

        levl_id = Select(driver.find_element(By.ID, "levl_id"))
        type_id = Select(driver.find_element(By.ID, "tprt_id"))

        #set path for student using their name
        path = "students/" + fullname[itr]
        build_path(path)
        build_files(path, driver, fullname[itr])

        driver.find_element(By.LINK_TEXT, "Faculty Services").click()
        itr += 1

if auth == "student":
#--WebDriver tends to fail here, maybe try rebooting the WebDriver?--
#------------WebDriver sometimes fail to grab name here--------------
    #grab person's name to use for file name
    name = driver.find_elements(By.XPATH, "//td[@class='pldefault']")
    for i in name:
        if i.text[0:7] == "Welcome":
            fullname.append(i.text.split(",")[1].replace(" ", "").replace(".", ""))
#------------WebDriver sometimes fail to grab name here--------------
#--------NoSuchElement Error Happens Here, WebDriver failure?--------
    #crawl webpage to academic transcript
    driver.find_element(By.LINK_TEXT, "Student").click()
    driver.find_element(By.LINK_TEXT, "Student Records").click()
#--------NoSuchElement Error Happens Here, WebDriver failure?--------
#--WebDriver tends to fail here, maybe try rebooting the WebDriver?--

    #crawl webpage to academic transcript
    driver.find_element(By.LINK_TEXT, "Academic Transcript").click()

    levl_id = Select(driver.find_element(By.ID, "levl_id"))
    type_id = Select(driver.find_element(By.ID, "type_id"))

    path = "students/" + fullname[0]
    build_path(path)
    build_files(path, driver, fullname[0])

#Web driver is no longer needed
driver.quit()
preprocess.main(fullname)
