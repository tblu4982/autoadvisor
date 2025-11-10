from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import NoSuchElementException, SessionNotCreatedException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import tkinter  # GUI
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
from datetime import datetime, timedelta
import os  # used for read/write to files
import re  # used for regular expressions
import sys  # used to stop execution under certain circumstances
import preprocess
from collections import defaultdict
from openpyxl import Workbook, load_workbook
import pandas as pd
import ast

VERSION = "2.5.0"
COURSE_IN_PROGRESS = re.compile(r"Course\(s\) in progress")

def main(vnums, thread):
    def build_path(path, fullname):
        # create folder using student name
        # check if name has been established
        if bool(fullname):
            # search to see if student folder already exists
            if not os.path.exists(path):
                os.makedirs(path)
                if os.path.exists(path):
                    print(thread + " -> Path successfully created!")
                else:
                    print(thread + " -> Failed to create path!")
            else:
                print(thread + " -> Path already exists!")
        else:
            sys.exit(thread + " -> An unexpected error has occurred: Unable to locate student's name!")

    def build_files(path, driver, fullname):
        # scrape transcript for courses and semesters
        try:
            wait.until(EC.visibility_of_all_elements_located((By.XPATH, "//table")))
            data = driver.find_elements(By.XPATH, "//table")

            courses = []
            semesters = []
            # boolean flags to distinguish semesters
            sem_start = sem_end = False
            # used to determine if we reached current/future semesters
            is_curr = False
            # used to hold course as we build it from data
            proto = ""
            course_marker = False
            
            output = []

            for i in data:
                output += i.text.split('\n')

            for line in output:
                # Find each semester as we iterate through scraped data
                # signifies the start of courses in progress
                if COURSE_IN_PROGRESS.search(line) and not is_curr:
                    is_curr = True
                if re.search("Subject Course Level Title Grade Credit Hours Quality Points Start and End Dates R|Subject Course Title Grade Credit hours Quality points R|Subject Course Level Title Credit Hours Start and End Dates|Subject Course Campus Level Title Credit Hours Start and End Dates", line):
                    if sem_start:
                        # lets us know that we've reached the end of a previous semester
                        sem_end = True
                    # signifies the start of a new semester
                    sem_start = True
                # if we are in a semester, find the courses
                elif sem_start:
                    # add marker to course array to signify semesters
                    if sem_end:
                        courses.append("-")
                        sem_end = False
                    # append credits to course
                    if course_marker:
                        if re.search("[0-9]+.[0-9]{3}", line):
                            course_marker = False
                            for i in line.replace("\n", " ").split(" "):
                                proto.append(i)
                            course = proto.copy()
                            # pop empty indices at beginning of course
                            while not bool(course[0]):
                                course.pop(0)
                            # remove 'U' from courses
                            if course[2] == 'U':
                                course.pop(2)
                            # remove unnecessary indices at end of course
                            try:
                                while re.search("[0-9]+.[0-9]{3}" ,course[-2]):
                                    course.pop()
                            except IndexError as e:
                                #print(e)
                                #print(course)
                                #print(line)
                                sys.exit()
                            # append course to array and reset proto
                            courses.append(course)
                            #print(proto)
                            proto.clear()
                        # check if we reached current/future semesters
                        else:
                            for i in line.replace("\n", " ").split(" "):
                                proto.append(i)
                    # find a course and store it to the holder array
                    elif re.search("[A-Z]{4}", line):
                        proto = line.replace("\n", " ").split(" ")
                        # Extra code to catch transfer courses and current semester courses, since their contents are stored in one line
                        if re.search("0.000$", line) or is_curr:
                            course = proto.copy()
                            # Add 'inprog' to current semester courses as they have no letter grade
                            if is_curr:
                                course.insert(-1, "inprog")
                            # remove unnecessary indices at end of course
                            else:
                                course.pop()
                            courses.append(course)
                            proto.clear()
                        else:
                            course_marker = True

            # Append separator for final course structure
            courses.append("-")
            
            data = driver.find_elements(By.CSS_SELECTOR, ".sub-heading.period-padding.ng-binding")
            
            output = []

            for i in data:
                output += i.text.split('\n')
                
            for line in output:
                semesters.append(line)
                semesters.append('-')

            # print courses to file
            create_file_path(fullname , path, "/courses.txt", "courses", courses)
            # print semesters to file
            create_file_path(fullname , path, "/semesters.txt", "semesters", semesters)
            
            return True
        except TimeoutException:
            return False
            
    # method to create file to store data
    def create_file_path(fullname ,path, filename, file_type, array):
        # check if file directory exists
        if os.path.exists(path):
            # if folder exists, check if file already exists
            file_path = path + filename
            # if file exists, overwrite it
            if os.path.exists(file_path):
                # remove old file, then create new file
                print(thread + " -> Overwriting " + file_type + " file for " + fullname + "...")
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
                print(thread +" -> Creating " + file_type + " file for " + fullname + "...")
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
    username = os.getenv("BANNER_ID")
    password = os.getenv("BANNER_PW")

    config_file = "CSCI_2020_TRANSCRIPT.xlsx"

    # --------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------
    # May be relevant: SessionNotCreatedException
    try:
        #driver = webdriver.Edge(options = options)
        driver = webdriver.Edge()
        driver.get('https://login.vsu.edu')
    except SessionNotCreatedException:
        print("Microsoft Edge and Edge Webdriver are not the same version. Check the Edge Webdriver!")
    # --------ADD A CHECK TO SEE IF WEBDRIVER IS CURRENT VERSION--------

    wait = WebDriverWait(driver, 10)
        
    # attempt to log in
    wait.until(EC.visibility_of_element_located((By.ID, "input28")))
    uid = driver.find_element(By.ID, "input28")
    uid.send_keys(username)
    driver.find_element(By.CSS_SELECTOR, ".button").click()

    wait = WebDriverWait(driver, 5)

    try:
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Verify with something else")))
        driver.find_element(By.LINK_TEXT, "Verify with something else").click()
    except TimeoutException:
        pass

    wait = WebDriverWait(driver, 600)

    # Get authentication method from user
    auth_type = "Okta Push Notification"
    test = driver.find_elements(By.XPATH, "//div[@class = 'authenticator-row clearfix']")
    for t in test:
        button = t.find_element(By.CSS_SELECTOR, "a[data-se='button']")
        target = button.get_attribute("aria-label")
        match auth_type:
            case "Google Authenticator":
                if(target == "Select Google Authenticator."):
                    button.click()
                    break
            case "Okta 2FA Code":
                if(target == "Select to enter a code from the Okta Verify app."):
                    button.click()
                    break
            case "Okta Push Notification":
                if(target == "Select to get a push notification to the Okta Verify app."):
                    button.click()
                    break

    wait = WebDriverWait(driver, 10)

    try:
        wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@type = 'password']")))    
        pwd = driver.find_element(By.XPATH, "//input[@type = 'password']")
        pwd.send_keys(password)
        driver.find_element(By.CSS_SELECTOR, ".button").click()
    except TimeoutException:
        pass

    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@aria-label='launch app Banner Faculty Self Service 9']")))
    driver.find_element(By.XPATH, "//a[@aria-label='launch app Banner Faculty Self Service 9']").click()

    original_window = driver.current_window_handle

    for window_handle in driver.window_handles:
        if window_handle != original_window:
            driver.switch_to.window(window_handle)
            
    second_window = driver.current_window_handle

    pattern = re.compile(r"^campus-v2report-search-(\d{4})-(\d{2})-(\d{2})\.csv$")

    latest_file = None
    latest_date = None

    for fname in os.listdir():
        match = pattern.match(fname)
        if match:
            y, m, d = map(int, match.groups())
            file_date = datetime(y, m, d)
            if latest_date is None or file_date > latest_date:
                latest_date = file_date
                latest_file = fname
                
    df = pd.read_csv(latest_file, skiprows=2)

    required = ["Student ID", "Email"]

    subset = df[required]

    rename_map = {}
    for c in subset.columns:
        if "Student ID" in c:
            rename_map[c] = "VNumber"
        elif "Email" in c:
            rename_map[c] = "Email"
            
    subset = subset.rename(columns=rename_map)

    wait.until(EC.title_is("Faculty Services Dashboard"))
    driver.find_element(By.LINK_TEXT, "Advising Student Profile").click()

    # Used to hold student names
    fullname = []
    names = []
    error_vnums = []

    is_anonymous = False
    sem_flag = False

    base_dir = "advisors"
    now = datetime.now()
    timestamp = now.strftime("%b-%d-%Y-%H-%M-%S")
    three_min_ago = now - timedelta(minutes=3)

    recent_folders = []

    # Walk the directory tree under "advisors/"
    for root, dirs, files in os.walk(base_dir):
        for folder in dirs:
            try:
                # Parse timestamp-style folder names (e.g., "Nov-07-2025-18-12-30")
                folder_dt = datetime.strptime(folder, "%b-%d-%Y-%H-%M-%S")
                if folder_dt >= three_min_ago:
                    recent_folders.append((folder_dt, folder))
            except ValueError:
                # Skip folders that don't follow the timestamp pattern
                continue

    # Pick the most recent valid one if available
    if recent_folders:
        # Sort by datetime descending and pick the latest
        recent_folders.sort(reverse=True)
        timestamp = recent_folders[0][1]
        print(f"Using most recent folder: {timestamp}")
    else:
        print(f"No recent folder found — creating new one: {timestamp}")

    status = Tk()
    status_txt = StringVar()
    status_txt.set(f"(0/{len(vnums)})")
    status_msg = Label(status, textvariable = status_txt).pack()

    progress = Progressbar(status, orient = HORIZONTAL, length = 100, maximum = len(vnums), mode = 'determinate')

    def status_update(i):
        progress['value'] = i
        status_txt.set(f"{int(progress['value'])}/{len(vnums)} Students Processed")
        #status_msg = Label(status, text = f"{int(progress['value'])}/{len(vnums)} Students Processed").pack()
        progress.update_idletasks()
        progress.update()

    progress.pack(pady = 10)
    #status.mainloop()

    # NEW: collect students per advisor for consolidated write-once workbook
    advisor_dict = defaultdict(list)  # { advisor: [ {"name": <str>, "vnum": <str>} ] }

    itr = 0
    # crawl through banner until we get to student id
    index = 0
    first = True
    while index < len(vnums):
        status_update(index)
        advisor = ""
        # requests student id and uses it to get to transcript
        wait.until(EC.visibility_of_element_located((By.ID, "s2id_select2-term")))
        sid = driver.find_element(By.ID, "idSearchInput")
        action = ActionChains(driver).send_keys_to_element(sid, vnums[index]).perform()
        wait = WebDriverWait(driver, 1)
        # Additional crawling logic if we've reached the last vnum in list
        #if index == len(vnums):
            #driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']").submit()
            # gets student name
        #try:
        try:
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.search-result.name')))
        except TimeoutException:
            action = ActionChains(driver).send_keys_to_element(sid, vnums[index]).send_keys_to_element(sid, Keys.ENTER).perform()
            try:
                wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.search-result.name')))
            except TimeoutException:
                print(f"{thread} -> Error! {vnums[index]} could not be found!")
                error_vnums.append([vnums.pop(index), "Invalid V-Number"])
                continue
        wait = WebDriverWait(driver, 10)
        name = driver.find_element(By.CSS_SELECTOR, '.search-result.name')
        if name != "No match found.":
            names.append(name.text.split(" "))
            if names[0] == "Mr":
                names.pop(0)
            fullname.append(name.text.replace(" ", "").replace(".", ""))
            wait.until(EC.element_to_be_clickable((By.ID, "term-go"))).click()
        else:
            print(f"{thread} -> Error! {vnums[index]} could not be found!")
            error_vnums.append(vnums.pop(index))
            continue

        # crawl webpage to academic transcript
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Academic Transcript")))
        try:
            advisor_listing = driver.find_element(By.CSS_SELECTOR, ".facultyLinkClass:nth-child(2)").text.split(" ")
            advisor = advisor_listing[-1]
        except NoSuchElementException:
            pass
        driver.find_element(By.LINK_TEXT, "Academic Transcript").click()
        
        for window_handle in driver.window_handles:
            if window_handle != original_window and window_handle != second_window:
                driver.switch_to.window(window_handle)
                   
        wait.until(EC.visibility_of_element_located((By.ID, "transcriptLevelSelection")))
        driver.find_element(By.ID, "transcriptLevelSelection").click()
        wait.until(EC.visibility_of_element_located((By.XPATH, "//li[@id='ui-select-choices-row-1-']/div/div")))
        driver.find_element(By.XPATH, "//li[@id='ui-select-choices-row-1-']/div/div").click()
        wait.until(EC.visibility_of_element_located((By.ID, "transcriptTypeSelection")))
        driver.find_element(By.ID, "transcriptTypeSelection").click()
        wait.until(EC.visibility_of_element_located((By.XPATH, "//li[@id='ui-select-choices-row-2-']/div/div")))
        driver.find_element(By.XPATH, "//li[@id='ui-select-choices-row-2-']/div/div").click()
        wait.until(EC.visibility_of_element_located((By.XPATH, "//button[contains(.,'Submit')]")))
        driver.find_element(By.XPATH, "//button[contains(.,'Submit')]").click()

        if not bool(advisor):
            advisor = "No Advisor Listed"
        """if not os.path.exists(os.path.join("advisors", advisor)):
            os.makedirs(os.path.join("advisors", advisor))
            print("Created new advisor directory for " + advisor)"""
            
        print(thread + " -> Adding " + fullname[index] + " to advisor " + advisor)
        print(thread + " -> (student " + str(index + 1) + " of " + str(len(vnums)) + ")")

        # set path for student using their name
        if is_anonymous:
            path = "advisors/" + timestamp + "/" + advisor + "/" + vnums[index].strip() + '/' + config_file.split('/')[-1].split('.')[0]
            build_path(path, vnums[index].strip())
            if (not build_files(path, driver, vnums[index].strip())):
                error_vnums.append([vnums.pop(index), "Transcript missing or unable to parse"])
                driver.close()
                driver.switch_to.window(second_window)
                driver.find_element(By.LINK_TEXT, "Advisee Search").click()
                fullname.pop(index)
                continue

            # record successful student for consolidated workbook (anonymous -> use V-number as name)
            advisor_dict[advisor].append({
                "name": vnums[index].strip(),
                "vnum": vnums[index].strip()
            })

        else:
            path = "advisors/" + timestamp + "/" + advisor + "/" + fullname[index].strip() + '/' + config_file.split('/')[-1].split('.')[0]
            build_path(path, fullname[index].strip())
            if (not build_files(path, driver, fullname[index].strip())):
                error_vnums.append([vnums.pop(index), "Transcript missing or unable to parse"])
                driver.close()
                driver.switch_to.window(second_window)
                driver.find_element(By.LINK_TEXT, "Advisee Search").click()
                fullname.pop(index)
                continue

            advisor_dict[advisor].append({
                "name": fullname[index].strip(),
                "vnum": vnums[index].strip()
            })

        driver.close()
        driver.switch_to.window(second_window)
        driver.find_element(By.LINK_TEXT, "Advisee Search").click()
        itr += 1
        index += 1

    # Web driver is no longer needed
    driver.quit()
    status.quit()

    # --------------CHANGE FUNCTION CALLS TO OTHER SCRIPTS SO IT READS ADVISOR FROM DIRECTORY------------

    if is_anonymous:
        preprocess.main(vnums, config_file, vnums, names, sem_flag, timestamp)
    else:
        preprocess.main(fullname, config_file, vnums, names, sem_flag, timestamp)

    # build consolidated student_list.xlsx under advisors/<timestamp>/ ----
    
    def write_consolidated_student_list(root_dir, advisor_dict, df, thread="Main"):
        """
        Create or append to a single workbook (student_list.xlsx) with one sheet per advisor.
        Each sheet has rows: Name | V-Number | Email.
        If the workbook already exists, advisor sheets are appended to or updated.
        """
        out_path = os.path.join(root_dir, "student_list_" + thread + ".xlsx")
        os.makedirs(root_dir, exist_ok=True)

        # If file exists, load it — otherwise create new workbook
        if os.path.exists(out_path):
            wb = load_workbook(out_path)
            print(f"{thread} -> Appending to existing workbook: {out_path}")
        else:
            wb = Workbook()
            print(f"{thread} -> Creating new workbook: {out_path}")

        # Create lookup dict for quick email access
        email_map = dict(zip(df["VNumber"], df["Email"]))

        advisors_in_order = list(advisor_dict.keys()) or ["No Advisor Listed"]

        for adv in advisors_in_order:
            # Reuse or create sheet for each advisor
            if adv[:31] in wb.sheetnames:
                ws = wb[adv[:31]]
            else:
                ws = wb.create_sheet(title=adv[:31])
                ws.append(["Name", "V-Number", "Email"])  # add headers if new sheet

            for s in advisor_dict.get(adv, []):
                vnum = s.get("vnum", "")
                email = email_map.get(vnum, "")
                ws.append([s.get("name", ""), vnum, email])

            # Adjust column width
            if advisor_dict.get(adv):
                longest_name = max((len(s["name"]) for s in advisor_dict[adv]), default=4)
                ws.column_dimensions["A"].width = longest_name + 4

        # Remove the default "Sheet" if it’s empty and no longer needed
        if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1 and not wb["Sheet"].max_row > 1:
            del wb["Sheet"]

        wb.save(out_path)
        print(f"{thread} -> Wrote consolidated workbook: {out_path}")

    root_dir = os.path.join("advisors", timestamp)
    write_consolidated_student_list(root_dir, advisor_dict, subset)

    if len(error_vnums) > 0:
        print(thread + ' -> AutoAdvisor has encountered an issue parsing some V-Numbers. Please check the error file', file=sys.stderr)
        f2 = "advisors/" + timestamp + "/error_vnums.txt"
        with open(f2, 'w') as outfile:
            outfile.write("AutoAdvisor has encountered an issue parsing the following V-Number(s):\n")
            for vnum in error_vnums:
                outfile.write(f"{vnum}\n")

    print(thread + ' -> Program complete! Check files for advisory report(s).')
    
if __name__ == "__main__":
    vnums = ast.literal_eval(sys.argv[1])
    tname = sys.argv[2]
    main(vnums, tname)