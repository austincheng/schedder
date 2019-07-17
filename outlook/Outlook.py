from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.common.keys import Keys
import os
import time
from selenium.webdriver.common.action_chains import ActionChains
import re
import win32clipboard
import json

class Outlook:
	def __init__(self, users, rooms, start, end, meeting_length, email):
		self.users = users
		self.rooms = rooms
		self.start = start
		self.end = end
		self.meeting_length = meeting_length
		self.all = users + rooms
		self.email = email

	def get_availability(self):
		driver = webdriver.Chrome()
		driver.get('https://developer.microsoft.com/en-us/graph/graph-explorer')

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ms-signin-button")))
		driver.find_element_by_id('ms-signin-button').click()

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "i0116")))
		driver.find_element_by_id('i0116').send_keys(self.email)

		# May need to do additional steps based on login

		time.sleep(0.5)

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "idSIButton9")))
		driver.find_element_by_id('idSIButton9').click()

		time.sleep(1)

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "idSIButton9")))
		driver.find_element_by_id('idSIButton9').click()

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"explorer-main\"]/request-editors/div/ul/li[2]")))
		driver.find_element_by_xpath('//*[@id="explorer-main"]/request-editors/div/ul/li[2]').click()

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"headers-editor\"]/table/tbody/tr[2]/td[1]/input")))
		driver.find_element_by_xpath('//*[@id=\"headers-editor\"]/table/tbody/tr[2]/td[1]/input').send_keys('Prefer')

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"headers-editor\"]/table/tbody/tr[2]/td[2]/input")))
		driver.find_element_by_xpath('//*[@id=\"headers-editor\"]/table/tbody/tr[2]/td[2]/input').send_keys('outlook.timezone=\"Pacific Standard Time\"')

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"headers-editor\"]/table/tbody/tr[3]/td[1]/input")))
		driver.find_element_by_xpath('//*[@id=\"headers-editor\"]/table/tbody/tr[3]/td[1]/input').send_keys('Content-Type')

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"headers-editor\"]/table/tbody/tr[3]/td[2]/input")))
		driver.find_element_by_xpath('//*[@id=\"headers-editor\"]/table/tbody/tr[3]/td[2]/input').send_keys('application/json')

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"explorer-main\"]/request-editors/div/ul/li[1]")))
		driver.find_element_by_xpath('//*[@id=\"explorer-main\"]/request-editors/div/ul/li[1]').click()

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"graph-request-url\"]/input")))
		driver.find_element_by_xpath('//*[@id=\"graph-request-url\"]/input').clear()
		driver.find_element_by_xpath('//*[@id=\"graph-request-url\"]/input').send_keys('https://graph.microsoft.com/v1.0/me/calendar/getschedule')

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"post-body-editor\"]/div[2]/div")))
		schedules = '['
		for i in range(len(self.all) ):
			user = self.all[i]
			schedules += '\"' + user + '\"'
			if i != len(self.all) - 1:
				schedules += ', '
		schedules += ']'

		div = driver.find_element_by_xpath('//*[@id=\"post-body-editor\"]/div[2]/div')
		actions = ActionChains(driver)
		actions.move_to_element(div)
		actions.click(div)
		actions.send_keys('{\"Schedules\": ' + schedules + ', ' + \
		    '\"StartTime\": {' + \
		        '\"dateTime\": \"' + self.start + '\",' + \
		        '\"timeZone\": \"Pacific Standard Time\"' + \
		    '},' + \
		    '\"EndTime\": {' + \
		        '\"dateTime\": \"' + self.end + '\",' + \
		        '\"timeZone\": \"Pacific Standard Time\"' + \
		    '},' + \
		    '\"availabilityViewInterval\": \"' + str(self.meeting_length) + '\"' + \
		'}')
		actions.perform()

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"httpMethodSelect\"]/div/button")))
		driver.find_element_by_xpath('//*[@id=\"httpMethodSelect\"]/div/button').click()

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "httpMethodSelect-select-POST")))
		driver.find_element_by_id('httpMethodSelect-select-POST').click()

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "submitBtn")))
		driver.find_element_by_id('submitBtn').click()

		time.sleep(1.5)
		div = driver.find_element_by_xpath("//*[@id=\"jsonViewer\"]/div[2]/div/div[3]")
		actions = ActionChains(driver)
		actions.move_to_element(div)
		actions.click(div)
		# Copying does not work on Mac in Chrome
		actions.key_down(Keys.CONTROL) # Keys.COMMAND on Mac
		actions.send_keys("a")
		actions.send_keys("c")
		actions.key_up(Keys.CONTROL) # Keys.COMMAND on Mac
		actions.perform()

		win32clipboard.OpenClipboard()
		json_str = win32clipboard.GetClipboardData()
		win32clipboard.CloseClipboard()
		json_dict = json.loads(json_str)
		userTimes = []
		roomTimes = []
		for obj in json_dict['value']:
			if obj['scheduleId'] in self.users:
				userTimes.append((obj['scheduleId'], obj['availabilityView']))
			else:
				roomTimes.append((obj['scheduleId'], obj['availabilityView']))

		driver.close()
		return userTimes, roomTimes

	@staticmethod
	def create_event(start, end, location, emails, meeting_subject, meeting_body):
		driver = webdriver.Chrome()
		driver.get('https://developer.microsoft.com/en-us/graph/graph-explorer')

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ms-signin-button")))
		driver.find_element_by_id('ms-signin-button').click()

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "i0116")))
		driver.find_element_by_id('i0116').send_keys(emails[0]) # Login credentials

		# May need to do additional steps based on login

		time.sleep(0.5)

		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "idSIButton9")))
		driver.find_element_by_id('idSIButton9').click()

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
			attendees += Outlook.email_box(email)
			if i != len(email) - 1:
				attendees += ', '
		attendees += ']'

		actions = ActionChains(driver)
		actions.move_to_element(div)
		actions.click(div)
		actions.send_keys('{\"subject\": \"' + meeting_subject + '\",\"body\": {\"contentType\": \"HTML\",\"content\": \"' + meeting_body + '\"},' + \
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

	@staticmethod
	def email_box(email):
		return "{{\"emailAddress\": {{ \"address\":\"{}\"}},\"type\": \"required\"}}".format(email)
