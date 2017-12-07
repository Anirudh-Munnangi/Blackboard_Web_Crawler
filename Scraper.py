from math import ceil
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import datetime as dt
import pyexcel as pe
# Initializing selenium essentials
path_to_chromedr='C:\chromedriver_win32\chromedriver'
browser=webdriver.Chrome(executable_path=path_to_chromedr)

# Scraping Starts####################################################################################
url='https://lacscheduling.uc.edu/TracWeb40/Default.html'
browser.get(url)
username=browser.find_element_by_name("Username")
password=browser.find_element_by_name("Password")

# Providing Credentials##############################################################################
username.send_keys("****")
password.send_keys("****")
browser.find_element_by_xpath('//*[@id="flow"]/table/tbody/tr/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[2]/form/table/tbody/tr[3]/td[2]/p/a/span').click()

# Switching to consultant profile####################################################################
for i in range(1000):
	try:
		browser.find_element_by_link_text("Switch to Consultant profile").click()
	except NoSuchElementException:
		break

# Executing the Reports Functionality#################################################################
browser.execute_script("reports()")
division=browser.find_element_by_id("popupContainer")
iframe=division.find_element_by_tag_name('iframe')
browser.switch_to_frame(iframe)
form_id=browser.find_element_by_name("entryinc")
form_id.find_element_by_name("CategoryID").send_keys("Students By") # Choosing "Students" By Option
time.sleep(14)
sel=Select(form_id.find_element_by_id("ReportID"))
sel.select_by_value('rpt008b') # Choosing "Visits/Students by ??" By Option
time.sleep(22)
form_id.find_element_by_id("StudentsEmail_btn").click()
time.sleep(2)
sel=Select(form_id.find_element_by_id("Centers_sel"))
sel.select_by_value("19")# Choosing "Mass Center Study Tables" By Option
time.sleep(1)

# Deciding the FROM dates to go into the report########################################################
cur_date= dt.date.today()
cur_day=dt.date.weekday(cur_date)
if cur_day==0:
	#month_to_put=Months[cur_date.month]
	new_date=cur_date
	month_to_put=cur_date.month
elif cur_day==1:
	timedel=dt.timedelta(days=-1)
	new_date=cur_date+timedel
	#month_to_put=Months[new_date.month]
	month_to_put=new_date.month
elif cur_day==2:
	timedel=dt.timedelta(days=-2)
	new_date=cur_date+timedel
	#month_to_put=Months[new_date.month]
	month_to_put=new_date.month
elif cur_day==3:
	timedel=dt.timedelta(days=-3)
	new_date=cur_date+timedel
	#month_to_put=Months[new_date.month]
	month_to_put=new_date.month
elif cur_day==4:
	timedel=dt.timedelta(days=-4)
	new_date=cur_date+timedel
	#month_to_put=Months[new_date.month]
	month_to_put=new_date.month
elif cur_day==5:
	timedel=dt.timedelta(days=-5)
	new_date=cur_date+timedel
	#month_to_put=Months[new_date.month]
	month_to_put=new_date.month
else:
	timedel=dt.timedelta(days=-6)
	new_date=cur_date+timedel
	#month_to_put=Months[new_date.month]
	month_to_put=new_date.month

# Loading the date decided into the webpage##################################################################

# Loading the month
form_id.find_element_by_xpath('//*[@id="FromDate_entry"]').click()
time.sleep(2)
browser.find_element_by_id("monthSelect").click()
time.sleep(2)
monthstr="monthDiv_"+str(month_to_put-1)
browser.find_element_by_id(monthstr).click()
row_element_one=new_date+dt.timedelta(days=-7) # New change to have one week older dates in the results excel sheet
# Loading the day

# Identifying week of the month CRITICAL PART OF THE CODE AND LOGIC AHEAD
first_day =new_date.replace(day=1)
dom = new_date.day
adjusted_dom = dom + first_day.weekday()
week_number=int(ceil(adjusted_dom/7.0))

if week_number==1:
	list_of_tds=browser.find_elements_by_xpath('//*[@id="calendarDiv"]/div[7]/table/tbody/tr[1]/td')
elif week_number==2:
	list_of_tds=browser.find_elements_by_xpath('//*[@id="calendarDiv"]/div[7]/table/tbody/tr[2]/td')
elif week_number==3:
	list_of_tds=browser.find_elements_by_xpath('//*[@id="calendarDiv"]/div[7]/table/tbody/tr[3]/td')
elif week_number==4:
	list_of_tds=browser.find_elements_by_xpath('//*[@id="calendarDiv"]/div[7]/table/tbody/tr[4]/td')
else:
	list_of_tds=browser.find_elements_by_xpath('//*[@id="calendarDiv"]/div[7]/table/tbody/tr[5]/td')

list_of_tds[0].click()
# Year remains the same as Noel wont search at the time interval of December 24th to Jan 8th :P
time.sleep(1)

# Deciding the TO dates to go into the report########################################################
cur_date= dt.date.today()
cur_day=dt.date.weekday(cur_date)
if cur_day==0:
	#month_to_put=Months[cur_date.month]
	timedel=dt.timedelta(days=4)
	new_date=cur_date+timedel
	month_to_put=cur_date.month
elif cur_day==1:
	timedel=dt.timedelta(days=3)
	new_date=cur_date+timedel
	#month_to_put=Months[new_date.month]
	month_to_put=new_date.month
elif cur_day==2:
	timedel=dt.timedelta(days=2)
	new_date=cur_date+timedel
	#month_to_put=Months[new_date.month]
	month_to_put=new_date.month
elif cur_day==3:
	timedel=dt.timedelta(days=1)
	new_date=cur_date+timedel
	#month_to_put=Months[new_date.month]
	month_to_put=new_date.month
elif cur_day==4:
	new_date=cur_date
	#month_to_put=Months[new_date.month]
	month_to_put=new_date.month
elif cur_day==5:
	timedel=dt.timedelta(days=-1)
	new_date=cur_date+timedel
	#month_to_put=Months[new_date.month]
	month_to_put=new_date.month
else:
	timedel=dt.timedelta(days=-2)
	new_date=cur_date+timedel
	#month_to_put=Months[new_date.month]
	month_to_put=new_date.month

# Loading the month
form_id.find_element_by_xpath('//*[@id="ToDate_entry"]').click()
time.sleep(2)
browser.find_element_by_id("monthSelect").click()
time.sleep(2)
monthstr="monthDiv_"+str(month_to_put-1)
browser.find_element_by_id(monthstr).click()
row_element_two=new_date+dt.timedelta(days=-7) # New change to have one week older dates in the results excel sheet

# Loading the day

# Identifying week of the month CRITICAL PART OF THE CODE AND LOGIC AHEAD
first_day =new_date.replace(day=1)
dom = new_date.day
adjusted_dom = dom + first_day.weekday()
week_number=int(ceil(adjusted_dom/7.0))

if week_number==1:
	list_of_tds=browser.find_elements_by_xpath('//*[@id="calendarDiv"]/div[7]/table/tbody/tr[1]/td')
elif week_number==2:
	list_of_tds=browser.find_elements_by_xpath('//*[@id="calendarDiv"]/div[7]/table/tbody/tr[2]/td')
elif week_number==3:
	list_of_tds=browser.find_elements_by_xpath('//*[@id="calendarDiv"]/div[7]/table/tbody/tr[3]/td')
elif week_number==4:
	list_of_tds=browser.find_elements_by_xpath('//*[@id="calendarDiv"]/div[7]/table/tbody/tr[4]/td')
else:
	list_of_tds=browser.find_elements_by_xpath('//*[@id="calendarDiv"]/div[7]/table/tbody/tr[5]/td')

list_of_tds[-1].click()

# Clicking on GENERATE ###################################################################################
time.sleep(3)
browser.find_element_by_xpath('//*[@id="blockcell_left"]/table/tbody/tr[5]/td/p/a').click()

# Switching to Second Window #############################################################################
browser.switch_to_window(browser.window_handles[1])
values=browser.find_elements_by_tag_name('b')
visits=values[0].text
hours=values[1].text
print("visits=",visits)
print("hours=",hours)
sheet=pe.get_sheet(file_name="C:\Python_Scripts\TutorTrac_Scraper\Output.xlsx")
sheet.row += [row_element_one,row_element_two,visits,hours]
sheet.save_as("C:\Python_Scripts\TutorTrac_Scraper\Output.xlsx")
time.sleep(3)

# Closing the Web Browser #################################################################################
browser.quit()
