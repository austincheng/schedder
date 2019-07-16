from recurrent import RecurringEvent
import datetime
import sys
from outlook.Outlook import Outlook
from schedder.Schedder import Schedder

def json_time(date, time, room):
	return "{{date: \"{}\", time: \"{}\", room: \"{}\"}}".format(date, time, room)

meeting_length = 30
# Bot Info
message = sys.argv[1]
users = sys.argv[2:]
with open('emails.txt', 'w+') as emails:
	for user in users:
		emails.write(user + '\n')

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

outlook_api = Outlook(users, rooms, start_string, end_string, meeting_length)
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
		line += json_time(date, time, room_available[i][0])
	if i != len(available) - 1:
		line += ", "
	line += '\n'
	dates_file.write(line)
dates_file.write(']')
dates_file.close()

schedder = Schedder("file:///C:/Users/aucheng/Desktop/hackathon/popup/index.html")
schedder.run()