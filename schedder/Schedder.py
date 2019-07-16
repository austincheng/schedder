from cefpython3 import cefpython as cef
import sys
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import datetime
import time

class Schedder:
	def __init__(self, file):
		self.file = file

	def run(self):
		def py_function(date, time_str, location):
			def email_box(email):
				return "{{\"emailAddress\": {{ \"address\":\"{}\"}},\"type\": \"required\"}}".format(email)

			def create_event(start, end, location, emails):
				driver = webdriver.Chrome()
				driver.get('https://developer.microsoft.com/en-us/graph/graph-explorer')

				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ms-signin-button")))
				driver.find_element_by_id('ms-signin-button').click()

				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "i0116")))
				driver.find_element_by_id('i0116').send_keys('aucheng@cisco.com')

				time.sleep(0.5)

				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "idSIButton9")))
				driver.find_element_by_id('idSIButton9').click()

				"""
				If not on Cisco VPN
				time.sleep(2)
				os.system('java Hackathon')
				"""

				time.sleep(1)

				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "idSIButton9")))
				driver.find_element_by_id('idSIButton9').click()

				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"explorer-main\"]/request-editors/div/ul/li[2]")))
				driver.find_element_by_xpath('//*[@id="explorer-main"]/request-editors/div/ul/li[2]').click()

				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"headers-editor\"]/table/tbody/tr[2]/td[1]/input")))
				driver.find_element_by_xpath('//*[@id=\"headers-editor\"]/table/tbody/tr[2]/td[1]/input').send_keys('Content-Type')

				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"headers-editor\"]/table/tbody/tr[2]/td[2]/input")))
				driver.find_element_by_xpath('//*[@id=\"headers-editor\"]/table/tbody/tr[2]/td[2]/input').send_keys('application/json')

				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"explorer-main\"]/request-editors/div/ul/li[1]")))
				driver.find_element_by_xpath('//*[@id=\"explorer-main\"]/request-editors/div/ul/li[1]').click()

				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"graph-request-url\"]/input")))
				driver.find_element_by_xpath('//*[@id=\"graph-request-url\"]/input').clear()
				driver.find_element_by_xpath('//*[@id=\"graph-request-url\"]/input').send_keys('https://graph.microsoft.com/v1.0/me/calendar/events')

				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"post-body-editor\"]/div[2]/div")))
				div = driver.find_element_by_xpath('//*[@id=\"post-body-editor\"]/div[2]/div')
				attendees = '['
				for i in range(len(emails)):
					email = emails[i]
					attendees += email_box(email)
					if i != len(email) - 1:
						attendees += ', '
				attendees += ']'

				actions = ActionChains(driver)
				actions.move_to_element(div)
				actions.click(div)
				actions.send_keys('{\"subject\": \"Hackathon Demo Meeting\",\"body\": {\"contentType\": \"HTML\",\"content\": \"Let\'s sync up for the Schedder Hackathon demo.\"},' + \
										'\"start\": {' + \
											'\"dateTime\": \"' + start + '\",' + \
											'\"timeZone\": \"Pacific Standard Time\"' + \
										'},' + \
										'\"end\": {' + \
											'\"dateTime\": \"' + end + '\",' + \
											'\"timeZone\": \"Pacific Standard Time\"' + \
										'},' + \
										'\"location\":{' + \
											'\"displayName\":\"' + location + '\"' + \
										'},' + \
										'\"attendees\":' + attendees + \
									'}')
				actions.perform()

				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"httpMethodSelect\"]/div/button")))
				driver.find_element_by_xpath('//*[@id=\"httpMethodSelect\"]/div/button').click()

				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "httpMethodSelect-select-POST")))
				driver.find_element_by_id('httpMethodSelect-select-POST').click()

				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "submitBtn")))
				driver.find_element_by_id('submitBtn').click()

				driver.close()

			start = datetime.datetime.strptime(date + " " + time_str, '%Y-%m-%d %H:%M:%S')
			end = start + datetime.timedelta(minutes=30)

			start_string = str(start).replace(' ', 'T')[:19]
			end_string = str(end).replace(' ', 'T')[:19]

			email_file = open('emails.txt', 'r')
			emails = [email[:-1] for email in email_file.readlines()]

			create_event(start_string, end_string, location, emails)

		sys.excepthook = cef.ExceptHook  # To shutdown all CEF processes on error
		cef.Initialize()
		browser = cef.CreateBrowserSync(url=self.file,
		                      window_title="Schedder")	
		bindings = cef.JavascriptBindings()
		bindings.SetFunction("py_function", py_function)
		browser.SetJavascriptBindings(bindings)
		cef.MessageLoop()
		del browser
		cef.Shutdown()