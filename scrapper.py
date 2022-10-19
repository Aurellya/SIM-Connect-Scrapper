
# scrap the SIM Connect timetable and convert it into xlsx file

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

# functions
def welcome_message():
	print("\n\n###############################################################")
	print("   ____  __  _  _     ___  __   __ _  __ _  ____  ___  ____ ")
	print("  / ___)(  )( \/ )   / __)/  \ (  ( \(  ( \(  __)/ __)(_  _)")
	print("  \___ \ )( / \/ \  ( (__(  O )/    //    / ) _)( (__   )(  ")
	print("  (____/(__)\_)(_/   \___)\__/ \_)__)\_)__)(____)\___) (__) ")
	print("\n      WELCOME TO SIM CONNECT TIMETABLE SCRAPPER PROGRAM")
	print("		       created by: Aurellya")
	print("###############################################################")


def closing_message():
	print("[SIM Connect] Processes completed successfully!")
	print("[SIM Connect] It took " + str(round(now_time - start_time, 1)) + " seconds to complete.\n")

	print(" ___ _ __  _ _   __  _  _ ___ __   ")
	print("| __| |  \| | |/' _/| || | __| _\  ")
	print("| _|| | | ' | |`._`.| >< | _|| v | ")
	print("|_| |_|_|\__|_||___/|_||_|___|__/  \n\n")


def loginAs():
	condition = True
	while condition:
		print("\nLogin as: ")
		print("1. Apply to Study")
		print("2. Student/Associate Lecturer/Alumni")
		print("3. Staff")
		print("4. Recruitment Agent")
		print("5. Apply to Teach")
	
		input1 = input("\n>> ")

		if input1 == "1" or input1 == "2" or input1 == "3" or input1 == "4" or input1 == "5":
			condition = False

	if input1 == 1:
		return "Applicant" 
	elif input1 == 2:
		return "Student"
	elif input1 == 3:
		return "Staff"
	elif input1 == 4:
		return "RecruitmentAgent"
	elif input1 == 5:
		return "ApplyToTeach"
	else: 
		return "Student"


def open_Web():
	# opens the SIM Connect link
	print("[SIM Connect] Opening Sim Connect...")
	print("[SIM Connect] This may take a while depending on your connection...")
	time.sleep(3) 
	url = "https://simconnect.simge.edu.sg/psp/paprd/EMPLOYEE/HRMS/s/WEBLIB_EOPPB.ISCRIPT1.FieldFormula.Iscript_SM_Redirect?cmd=login"
	driver.get(url)

	# gives time for the initial loading of the website to prevent errors
	WebDriverWait(driver, 10).until(EC.presence_of_element_located(("id","login")))


def month_conversion(i):
	switcher = {
		'Jan' : 1,
		'Feb' : 2,
		'Mar' : 3,
		'Apr' : 4,
		'May' : 5,
		'Jun' : 6,
		'Jul' : 7,
		'Aug' : 8,
		'Sep' : 9,
		'Oct' : 10,
		'Nov' : 11,
		'Dec' : 12
	}

	return switcher.get(i, "Invalid")


def xlsx_writer(final_schedule, date, month_int, subject, tutorialOrLectureGroup, subjects, typeOfClass, timeSpan, location, instructors):
	print("[SIM Connect] Exporting to XLSX file...")

	d = {'Dates': final_schedule, 'Date': date, 'Month Int': month_int ,'Subject': subjects, 'Group': tutorialOrLectureGroup, 'Class Type': typeOfClass, 'Time': timeSpan, 'Location': location, 'Instructor': instructors}
	df = pd.DataFrame(data = d)
	result = df.sort_values(by =['Month Int', 'Date'])
	result = result.drop(columns=['Month Int' ,'Date'])
	writer = pd.ExcelWriter('schedule.xlsx', engine = 'xlsxwriter')
	result.to_excel(writer, sheet_name = 'Timetable', index = False)
	writer.save()

###############################################################################################################################

# program starts (to calculate the time)
start_time = time.time() 

# print welcome message	
welcome_message() 	

# login as			
loginAs = loginAs()					

# authentication
print("\n[SIM Connect] Enter your ID: ", end = '')
my_username = input()
print("[SIM Connect] Enter your password: ", end = '')
my_password = input()

# number of weeks parameter
print("\nHow many Weeks?")
weeks = int(input(">> "))			

# accesses chromedriver
print("\n[SIM Connect] Opening Chrome web browser...")
driver = webdriver.Chrome()			

# to update chromedriver = 'brew cask upgrade chromedriver'

# opens the SIM Connect link
open_Web()							

# finds the username and password forms
user_type = driver.find_element("id", "User_Type")
username = driver.find_element("id", "userid")
password = driver.find_element("id", "pwd")
login_button = driver.find_element("id", "loginbutton")

# fills in the login form and submits
print("[SIM Connect] Entering credentials...")
time.sleep(2)
user_type.send_keys(loginAs)
username.send_keys(my_username)
password.send_keys(my_password)
login_button.click()
print("[SIM Connect] Logged in successfully!")

# steps to get to the weekly schedule page which contains the schedule
print("[SIM Connect] Moving to My Apps Page ...")
driver.find_element("link text", 'My Apps').click()

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'fldra_CO_EMPLOYEE_SELF_SERVICE')))

print("[SIM Connect] Moving to Self-Service Page ...")
driver.find_element("id", 'fldra_CO_EMPLOYEE_SELF_SERVICE').click()

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'fldra_HCCC_ENROLLMENT')))

print("[SIM Connect] Moving to Enrollment Page ...")
driver.find_element("id", 'fldra_HCCC_ENROLLMENT').click()

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'ptifrmtemplate')))

print("[SIM Connect] Moving to Weekly Timetable Page ...")
driver.find_element("link text",'My Timetable (Weekly View)').click()

###############################################################################################################################

# passes the html page to the beautiful soup parser
simConnect_html = BeautifulSoup(driver.page_source, "html.parser")
driver.implicitly_wait(8)

# takes the src link to get to the timetable sourcecode  
timetable_data = simConnect_html.find(id="ptifrmtgtframe")
timetable_data_src = timetable_data['src']
driver.get(timetable_data_src)
print("[SIM Connect] You are in the Weekly Timetable Page ...")

# initializes arrays to store the schedule
final_schedule = []		# list of dates and months that has scheduled class 
subjects = []
tutorialOrLectureGroup = []
typeOfClass = []				
timeSpan = []
location = [] 
instructors = []

# for sorting purpose	
date = [] 				# list of dates that has scheduled class 
month = [] 				# list of months that has scheduled class 
month_int = [] 			# list of months that has scheduled class in integer form 

index = 0
while index < weeks:
	# passes the html page to the beautiful soup parser
	print("[SIM Connect] Reading the page...")
	simConnect_html = BeautifulSoup(driver.page_source, "html.parser")

	calendar = [] 			# list of all dates and months 
	calendar_date = [] 		# list of all dates
	calendar_month = [] 	# list of all months

	# dates entry
	dates_list = simConnect_html.findAll("th",{"class":"SSSWEEKLYDAYBACKGROUND"})
	for date1 in dates_list:
		entry = date1.text.strip()
		entry = entry.split("\n")
		entry1 = entry[1].split(" ")
		calendar_month.append(entry1[1])
		calendar_date.append(int(entry1[0]))
		new_entry = entry[0] + " " + entry[1]
		calendar.append(new_entry)

	# finding schedule's parent
	tags = simConnect_html.find("table",{"id":"WEEKLY_SCHED_HTMLAREA"})
	tags = tags.find("tbody")
	tags = tags.findAll("tr")

	tr_tag = [] 		# list of tr

	for tag in tags:
		tr_tag.append(tag)

	tr_tag_final =[]	# list of tr without the first index of tr_tag	
	td_tag = [] 		# list of td in each tr

	i = 1
	while i < len(tr_tag):
		tr_tag_final.append(tr_tag[i])
		other_td = tr_tag_final[i-1].findAll("td")
		temporary = []
		for other in other_td:
			temporary.append(other)
		td_tag.append(temporary)
		i = i + 1

	final_keywords_index = [] 	# index of "SSSWEEKLYBACKGROUND" where the schedule is

	j = 0
	while j < len(td_tag):
		temporary = []
		temporary.clear()
		classes_in_j_index = td_tag[j]
		get_classes = str(classes_in_j_index).split('\"')

		for get_class in get_classes:
			if get_class[0:9] == "SSSWEEKLY" and get_class[0:13] != "SSSWEEKLYTIME":
				temporary.append(get_class)
				
		k = 0
		while k < len(temporary):
			if temporary[k] == "SSSWEEKLYBACKGROUND":
				final_keywords_index.append(k)
			k = k + 1
		j = j + 1

	for f in final_keywords_index:
		final_schedule.append(calendar[int(f)])
		date.append(calendar_date[int(f)])
		month.append(calendar_month[int(f)])


	# list of classes details includes subject name, class type, etc.
	subject_list = simConnect_html.findAll("span",{"class":"SSSTEXTWEEKLY"})

	# get details
	for subject in subject_list:
		entry1 = str(subject.contents[0]) # subject name (subject name and lecture or tutor group)
		en1 = entry1[0:-6]
		en2 = entry1[-3:]
		entry2 = str(subject.contents[2]) # type of class
		entry3 = str(subject.contents[4]) # time
		entry4 = str(subject.contents[6]) # location
		try:
			entry5 = str(subject.contents[10]) # instructor name
		except:
			entry5 = " "
		
		subjects.append(en1)
		tutorialOrLectureGroup.append(en2)
		typeOfClass.append(entry2)
		timeSpan.append(entry3)
		location.append(entry4)
		instructors.append(entry5)

	# going to the next page which has the timetable for the following week
	next = driver.find_element("id", "DERIVED_CLASS_S_SSR_NEXT_WEEK")
	next.click()
	time.sleep(3)
	index = index + 1

# convert month from string type to integer
for m in month:
	month_converted = month_conversion(m)
	month_int.append(month_converted)

###############################################################################################################################

# closes the web browser
print("[SIM Connect] Information retrieval complete.")
driver.close()

###############################################################################################################################

#  export data to excel file (.xlsx) using pandas library
xlsx_writer(final_schedule, date, month_int, subjects, tutorialOrLectureGroup, subjects, typeOfClass, timeSpan, location, instructors)

###############################################################################################################################

# print time taken in console
now_time = time.time()
closing_message()




