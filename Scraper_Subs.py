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

# Loading the Dates
time.sleep(1)
form_id.find_element_by_xpath('//*[@id="dateRanges"]/option[11]').click()
time.sleep(1)
form_id.find_element_by_xpath('//*[@id="fieldName"]').click()
time.sleep(2)
form_id.find_element_by_xpath('//*[@id="fieldName"]/option[26]').click()
time.sleep(1)
###
form_id.find_element_by_xpath('//*[@id="fieldName"]').click()
time.sleep(2)
row_element_one=form_id.find_element_by_id("FromDate_entry").get_attribute('value')
row_element_two=form_id.find_element_by_id("ToDate_entry").get_attribute('value')
# Clicking on GENERATE ###################################################################################
time.sleep(3)
browser.find_element_by_xpath('//*[@id="blockcell_left"]/table/tbody/tr[5]/td/p/a').click()

# Switching to Second Window #############################################################################
browser.switch_to_window(browser.window_handles[1])
list_of_bolds=browser.find_elements_by_class_name('bold') # Returns bigfont bold also
# The below code weeds out the big font bolds and just keeps the normal bolds
# It is observed  that at first there are 3 big font bolds to be ignored and the rest are alternate
# Keep running until evidence as above is verified
length=len(list_of_bolds)
newlist=list_of_bolds[3:length:2]
sheet=pe.get_sheet(file_name="C:\Python_Scripts\TutorTrac_Scraper\Output.xlsx")
#Uncomment tv/th related lines for totals etry
#tv=0
#th=0
for value in newlist:
	cur_parent=value.find_element_by_xpath('..')
	list_of_tds=cur_parent.find_elements_by_tag_name('td')
	row_element_three=list_of_tds[0].text
	row_element_three=row_element_three[0:-7] # Cropping the collected string name
	if not row_element_three:
		row_element_three="Unlisted"
	row_element_four=list_of_tds[3].text
	#tv+=float(row_element_four)
	row_element_five=list_of_tds[4].text
	#th+=float(row_element_five)
	sheet.row += [row_element_one,row_element_two,row_element_three,row_element_four,row_element_five]

#sheet.row+=[row_element_one,row_element_two,"All Sessions",tv,th]
sheet.save_as("C:\Python_Scripts\TutorTrac_Scraper\Output.xlsx")
time.sleep(3)

# Closing the Web Browser #################################################################################
browser.quit()
