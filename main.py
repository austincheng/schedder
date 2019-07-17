from recurrent import RecurringEvent
import datetime
import sys
import os
from outlook.Outlook import Outlook
from schedder.Schedder import Schedder

# Changeable meeting parameters
meeting_length = 30 # Can change based on length of meeting in minutes
meeting_subject = "Hackathon Demo Meeting"
meeting_body = "Let\'s sync up for the Schedder hackathon demo."

message = sys.argv[1]
# First email in list is the login credentials to Microsoft API, and therefore also the sender of the invite
email = sys.argv[2]
users = sys.argv[2:]

# Can add additional rooms
rooms = ['CONF_100242@cisco.com', 'CONF_70358@cisco.com']
room_map = {'CONF_100242@cisco.com': 'SJC09-1-FURMAN',
			'CONF_70358@cisco.com': 'SJC09-1-CHIMNEY ROCK'}

r = RecurringEvent()
date = r.parse(message)
now = datetime.datetime.now()
if date is None:
	start = now
	end = start + datetime.timedelta(days=30)
else:
	if 'week' in message or 'month' in message or 'year' in message:
		# Get all meetings within business week
		weekday = date.weekday()
		if weekday == 5 or weekday == 6:
			# Saturday or Sunday consider next Monday
			start = date + datetime.timedelta(days=7 - weekday)
		else:
			# That week's Monday
			start = date - datetime.timedelta(days=weekday)
		end = start + datetime.timedelta(days=4)
	else:
		if (date.second == 0 and date.microsecond == 0) and (date.hour != 9):
			# Specific time
			start = date
			end = start + datetime.timedelta(minutes=meeting_length)
		else:
			# Get all meetings in that day
			start = datetime.datetime(date.year, date.month, date.day, 0)
			end = start + datetime.timedelta(days=1)
start_string = str(start).replace(' ', 'T')[:19]
end_string = str(end).replace(' ', 'T')[:19]

outlook_api = Outlook(users, rooms, start_string, end_string, meeting_length, email)
userTimes, roomTimes = outlook_api.get_availability()

len_1 = len(roomTimes[0][1]) == 1
unavailable_people = []
available = []
for i in range(len(roomTimes[0][1])):
	bad = False
	for user, availability in userTimes:	
		if len_1:
			if availability != '0':
				unavailable_people.append(user)
		else:
			if availability[i] != '0':
				bad = True
				break
	if not bad:
		available.append(i)

room_available = []
for i in available:
	rooms_available = []
	for room, availability in roomTimes:
		if availability[i] == '0':
			rooms_available.append(room_map[room])
	room_available.append(rooms_available)

dates_file = open('./popup/dates.js', 'w+')
dates_file.write('dates = [')
for i in range(len(available)):
	time = str(start + datetime.timedelta(minutes=meeting_length * available[i]))[:19]
	date_time = time.split(' ')
	date = date_time[0]
	time = date_time[1]
	line = ''
	if len(room_available[i]) == 0:
		continue
	else:
		line += "{{date: \"{}\", time: \"{}\", room: \"{}\"}}".format(date, time, room_available[i][0])
	if i != len(available) - 1:
		line += ", "
	line += '\n'
	dates_file.write(line)
dates_file.write(']')
dates_file.close()

local_path_to_folder = os.path.dirname(os.path.abspath(__file__))
full_path_to_index = "file:///" + local_path_to_folder + "/popup/index.html" 
schedder = Schedder(full_path_to_index, users, meeting_length, meeting_subject, meeting_body)
schedder.run()