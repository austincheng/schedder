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
from outlook.Outlook import Outlook
import datetime
import time

class Schedder:
	def __init__(self, file):
		self.file = file

	def run(self):
		sys.excepthook = cef.ExceptHook  # To shutdown all CEF processes on error
		cef.Initialize()
		browser = cef.CreateBrowserSync(url=self.file,
		                      window_title="Schedder")	
		bindings = cef.JavascriptBindings()
		bindings.SetFunction("py_function", Schedder.py_function)
		browser.SetJavascriptBindings(bindings)
		cef.MessageLoop()
		del browser
		cef.Shutdown()

	@staticmethod
	def py_function(date, time_str, location):
		start = datetime.datetime.strptime(date + " " + time_str, '%Y-%m-%d %H:%M:%S')
		end = start + datetime.timedelta(minutes=30)

		start_string = str(start).replace(' ', 'T')[:19]
		end_string = str(end).replace(' ', 'T')[:19]

		email_file = open('emails.txt', 'r')
		emails = [email[:-1] for email in email_file.readlines()]

		Outlook.create_event(start_string, end_string, location, emails)