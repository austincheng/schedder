from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.common.keys import Keys
import csv
import os
import time
from selenium.webdriver.common.action_chains import ActionChains
import re
import win32clipboard
import json

class Outlook:
	def __init__(self, users, rooms, start, end, meeting_length):
		self.users = users
		self.rooms = rooms
		self.start = start
		self.end = end
		self.meeting_length = meeting_length
		self.all = users + rooms

	def get_availability(self):
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
		actions.key_down(Keys.CONTROL)
		actions.send_keys("a")
		actions.send_keys("c")
		actions.key_up(Keys.CONTROL)
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
