from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import NoSuchElementException
import tkinter  # GUI
from tkinter import *
from tkinter import messagebox
from datetime import datetime
import os  # used for read/write to files
import re  # used for regular expressions
import sys  # used to stop execution under certain circumstances
import preprocess
import time


# Class to hold user login info for session
class credentials():
    # Opens a window and prompts user to select if they are student or advisor
    def greeting_window(self, greeting):
        portal = Tk()
        portal.geometry("200x70")
        portal.title("Auto Advisor")
        welcome = Label(portal, text=greeting)
        login = Button(portal, text="Log In", command=lambda: [portal.destroy(), self.click_login()])

        welcome.pack(side=TOP)
        login.pack(side=BOTTOM)
        portal.mainloop()

    # method to grab student login info
    def click_login(self):
        # opens a window to grab student login info
        login = Tk()
        login.geometry("200x90")
        login.title("Auto Advisor - Student")
        welcome = Label(login, text="Enter your V-Number and PIN:")
        Label(login, text='V-Number').grid(row=1)
        Label(login, text='PIN').grid(row=2)
        uid = Entry(login)
        pwd = Entry(login, show="*")
        # when submit button is clicked, it sends credentials to Banner Portal
        submit = Button(login, text='Submit', command=lambda: [self.set_credentials(uid, pwd), login.destroy()])
        back = Button(login, text='Return', command=lambda: [login.destroy(), self.greeting_window(self.greeting)])
        welcome.grid(row=0, columnspan=3)
        uid.grid(row=1, column=1, columnspan=2)
        pwd.grid(row=2, column=1, columnspan=2)
        submit.grid(row=3, column=1)
        back.grid(row=3, column=2)

    # sets user's login info
    def set_credentials(self, uid, pwd):
        self.username = uid.get()
        self.password = pwd.get()

    # returns login info
    def get_credentials(self):
        return self.username, self.password

    # warns the user that they have failed to log in 3 times
    def login_warning(self):
        warn = Tk()
        warn.geometry("200x90")
        warn.title("Warning!")
        warning = Label(
            text="Warning! You have failed to login 3 times now. Please ensure that you enter your information correctly.")
        back = Button(warn, text='Return', command=lambda: [warn.destroy(), self.greeting_window(self.greeting)])
        warning.pack(side=TOP)
        back.pack(side=TOP)
        warn.mainloop()

    # terminates the program if user fails to log in 4 times
    def login_timeout(self):
        terminate = Tk()
        terminate.geometry("200x90")
        terminate.title("Too Many Login Attempts")
        message = Label(
            text="You have attempted too many failed login attempts. To prevent account lockout, please try again later.")
        end = Button(terminate, text='Exit', command=lambda: [terminate.destroy(), driver.quit(),
                                                              sys.exit("Program Terminated: Too Many Login Attempts!")])
        message.pack(side=TOP)
        end.pack(side=TOP)
        terminate.mainloop()

    # method that prompts user to correct invalid login info
    def invalid_creds(self):
        self.greeting = 'Invalid login credentials!'
        self.login_count += 1
        if self.login_count < 3:
            self.greeting_window(self.greeting)
        elif self.login_count == 3:
            self.login_warning()
        else:
            self.login_timeout()

    # instantiates class
    def __init__(self):
        self.username = ""
        self.password = ""
        self.greeting = 'Welcome to Auto-Advisor!'
        self.greeting_window(self.greeting)
        self.login_count = 0


# method to return appropriate value for term dropdown
def get_timecode():
    # get current month and year
    month = datetime.now().month
    year = str(datetime.now().year)
    # change month to string
    if month >= 1:
        # set month to "01" if month is between 1 and 4
        if month >= 5:
            # set month to "05" if month is between 5 and 7
            if month >= 8:
                # set month to "08" if month is between 8 and 11
                if month == 12:
                    # change month to string if month is 12
                    month = "12"
                else:
                    month = "08"
            else:
                month = "05"
        else:
            month = "01"
    # return year and month to use as value for dropdown selector
    return year + month


def build_path(path):
    # create folder using student name
    # check if name has been established
    if bool(fullname):
        # search to see if student folder already exists
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
    print("Building course and semester files...")
    time.sleep(0.1)
    driver.find_element(By.XPATH, "//table[contains(@class, 'dataentrytable')]").submit()
    time.sleep(1)

    # scrape transcript for courses and semesters
    data = driver.find_elements(By.XPATH, "//table[@class = 'datadisplaytable']")

    courses = []
    semesters = []
    # boolean flags to distinguish semesters
    sem_start = sem_end = False
    # used to determine if we reached current/future semesters
    is_curr = False
    # used to hold course as we build it from data
    proto = ""
    output = ''

    for i in data:
        output = i.text.split('\n')

    for line in output:
        print(line)
        # Find each semester as we iterate through scraped data
        # signifies the start of courses in progress
        if re.search("COURSES IN PROGRESS", line) and not is_curr:
            is_curr = True
        if re.search("Spring", line) or re.search("Summer", line) or \
                re.search("Fall", line) or re.search("Winter", line) or \
                re.search("1:", line):
            if sem_start:
                # lets us know that we've reached the end of a previous semester
                sem_end = True
            # signifies the start of a new semester
            sem_start = True
            # store the current semester into semesters array
            if re.search("Spring [0-9]{4}|Summer [0-9]{4}|Fall [0-9]{4}|Winter [0-9]{4}", line):
                semesters.append(
                    re.search("Spring [0-9]{4}|Summer [0-9]{4}|Fall [0-9]{4}|Winter [0-9]{4}", line).group())
            # for semesters that are missing a term name
            else:
                semesters.append("UNKNOWN")
            semesters.append("-")
        # if we are in a semester, find the courses
        elif sem_start:
            # add marker to course array to signify semesters
            if sem_end:
                courses.append("-")
                sem_end = False
            # append credits to course
            if bool(proto):
                proto.append(line)
                if len(proto) > 6:
                    course = proto.copy()
                    # remove 'U' from courses
                    # check if we reached current/future semesters
                    if is_curr:
                        # add 'inprog' to current/future courses
                        course.insert(4, "inprog")
                        course.pop()
                    if course[2] == 'U':
                        course.pop(2)
                    course[4] = course[4].strip()
                    course.pop()
                    if re.search("0.000", course[-1]):
                        course.pop()
                    print(course)
                    # append course to array and reset proto
                    courses.append(course)
                    proto.clear()
            # find a course and store it to the holder array
            elif re.search("^[A-Z]{4}", line) and not re.search("-Top-", line):
                proto = line.replace("\n", " ").split(" ")

    # Append separator for final course structure
    courses.append("-")

    # print courses to file
    create_file_path(fullname, path, "/courses.txt", "courses", courses)
    # print semesters to file
    create_file_path(fullname, path, "/semesters.txt", "semesters", semesters)


# method to create file to store data
def create_file_path(fullname, path, filename, file_type, array):
    # check if file directory exists
    if os.path.exists(path):
        # if folder exists, check if file already exists
        file_path = path + filename
        # if file exists, overwrite it
        if os.path.exists(file_path):
            # remove old file, then create new file
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
            # create new file
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


# Gets user login credentials
user = credentials()
try:
    username, password = user.get_credentials()
# Exception Handling that closes program if tkinter box is closed prematurely
except AttributeError:
    sys.exit("Program Terminated")

if not bool(username) or not bool(password):
    sys.exit("Credentials not set")

# --------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------
# May be relevant: SessionNotCreatedException
options = Options()
options.add_argument("headless")
driver = webdriver.Safari()
driver.get('https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_WWWLogin')
# --------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------

# enter credentials into Banner
uid = driver.find_element(By.ID, "UserID")
pwd = driver.find_element(By.ID, "PIN")

uid.send_keys(username)
pwd.send_keys(password)

driver.find_element(By.XPATH, "//form").submit()

# Track if user is at landing page
home_url = "https://ssb-prod.ec.vsu.edu/BNPROD/twbkwbis.P_GenMenu?name=bmenu.P_MainMnu"
time.sleep(1)
url = driver.current_url.split("&")[0]
print(url)
# if we are not at expected landing page, then user hasn't logged in yet
while not home_url == url:
    # set new login credentials
    user.invalid_creds()
    # store old credentials
    username_old = username
    password_old = password
    # get new credentials
    username, password = user.get_credentials()
    # if old credentials match new credentials, terminate program
    if username_old == username and password_old == password:
        driver.quit()
        sys.exit(
            "Program Terminated: You have either manually closed the program, or entered the same invalid login credentials twice in a row.")
    # attempt to log in again
    uid = driver.find_element(By.ID, "UserID")
    pwd = driver.find_element(By.ID, "PIN")

    uid.send_keys(username)
    pwd.send_keys(password)

    driver.find_element(By.XPATH, "//form").submit()

    url = driver.current_url.split("&")[0]
# detect whether student or advisor is logged in by available links
auth = ''
# if this link is available, then advisor is logged in
try:
    driver.find_element(By.LINK_TEXT, "Faculty and Advisors")
    auth = 'advisor'
# if this link can't be found, then student is logged in
except NoSuchElementException:
    auth = 'student'

# Used to hold student V-Numbers
vnums = []
try:
    f1 = open("vnums.txt")
    for vnum in f1:
        vnums.append(vnum)
    f1.close()
except FileNotFoundError:
    pass
# Used to hold student names
fullname = []

itr = 0
if auth == "advisor":
    # crawl through banner until we get to student id
    driver.find_element(By.LINK_TEXT, "Faculty and Advisors").click()
    time.sleep(0.1)
    for vnum in vnums:
        driver.find_element(By.LINK_TEXT, "Student Information Menu").click()
        time.sleep(0.1)
        driver.find_element(By.LINK_TEXT, "ID Selection").click()
        time.sleep(0.1)
        # The following dropdown menu lacks an id
        # so to select for current semester, I noticed html code stored
        # dropdown values as dates, I used a method to grab current
        # month and year, and store it to a variable to use as value
        date = get_timecode()
        # We visit the 'Select Term' page only once, during the first time we crawl
        if vnum == vnums[0]:
            term = Select(driver.find_element(By.XPATH, "//select[@name='term']"))
            term.select_by_value(date)
            driver.find_element(By.XPATH, "//td[@class='dedefault']").submit()
            time.sleep(0.1)
        # requests student id and uses it to get to transcript
        sid = driver.find_element(By.ID, "Stu_ID")
        sid.send_keys(vnum)
        # Additional crawling logic if we've reached the last vnum in list
        if vnum == vnums[-1]:
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
            time.sleep(0.1)
            # gets student name
            name = driver.find_element(By.TAG_NAME, 'b')
            fullname.append(name.text.replace(" ", "").replace(".", ""))
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
            time.sleep(0.1)
        # For all other vnums that are not last in list
        else:
            name = driver.find_element(By.TAG_NAME, 'b')
            fullname.append(name.text.replace(" ", "").replace(".", ""))
            driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
            time.sleep(0.1)

        # crawl webpage to academic transcript
        driver.find_element(By.LINK_TEXT, "Academic Transcript").click()
        time.sleep(0.1)

        levl_id = Select(driver.find_element(By.ID, "levl_id"))
        type_id = Select(driver.find_element(By.ID, "tprt_id"))

        # set path for student using their name
        path = "students/" + fullname[itr]
        build_path(path)
        build_files(path, driver, fullname[itr])

        driver.find_element(By.LINK_TEXT, "Faculty Services").click()
        itr += 1

if auth == "student":
    # --WebDriver tends to fail here when 'headless' Option is enabled, maybe try rebooting the WebDriver?--
    # ------------WebDriver sometimes fail to grab name here--------------
    # grab person's name to use for file name
    j = 0
    while fullname == [] or fullname == ['']:
        if j > 25:
            sys.exit("Sytem Timeout: Failed to get name!")
        name = driver.find_elements(By.XPATH, "//td[@class='pldefault']")
        for i in name:
            print(i.text)
            if re.search("Welcome", i.text):
                fullname.append(i.text.split(",")[1].replace(" ", "").replace(".", ""))
        if j > 0:
            print("ERROR: Failed to get name! Retrying...")
        j += 1
    print("Name found: " + str(fullname))
    # ------------WebDriver sometimes fail to grab name here--------------
    # --------NoSuchElement Error Happens Here, WebDriver failure?--------
    # crawl webpage to academic transcript
    found_link = False
    i = 0
    while not found_link:
        if i > 10:
            sys.exit("System Timeout: Could not find link for 'Student'!")
        try:
            driver.find_element(By.LINK_TEXT, "Student").click()
            time.sleep(0.1)
            found_link = True
        except NoSuchElementException:
            pass
        if i > 0:
            print("ERROR: Could not find link 'Student'! Retrying...")
        i += 1
    driver.find_element(By.LINK_TEXT, "Student Records").click()
    time.sleep(0.1)
    # --------NoSuchElement Error Happens Here, WebDriver failure?--------
    # --WebDriver tends to fail here when 'headless' Option is enabled, maybe try rebooting the WebDriver?--

    # crawl webpage to academic transcript
    driver.find_element(By.LINK_TEXT, "Academic Transcript").click()
    time.sleep(0.1)

    levl_id = Select(driver.find_element(By.ID, "levl_id"))
    type_id = Select(driver.find_element(By.ID, "type_id"))

    path = "students/" + fullname[0]

    build_path(path)
    build_files(path, driver, fullname[0])

# Web driver is no longer needed
driver.quit()
preprocess.main(fullname)
