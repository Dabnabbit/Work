#!/bin/python
# UNCLASSIFIED
# mm_csver_v3.py
# Grabs information from MatterMost and presents it in readable format - in the event that you made things hard for yourself by not organizing MatterMost by date...
# V1 by SSgt Bentz.  Updated to V3 by TSgt Hayes

import os, subprocess
import requests
import pycurl
import time, datetime
from StringIO import StringIO
import json



#// Debug leading string to indicate message is coming from CSVer
mm_CSVer_debugPreStr = "[=-MMCSVer-=]:"
#// Debug error string for failed commands
mm_CSVer_errorPreStr = '!!!ERROR!!!'

#// Declare dictionaries for storing teams, channel/post, and user databases
mm_db = {}
mm_users = {}
#// Declare list for storing header data, generated from cookie gathering
mm_header = []

#// Declare static MatterMost URL library
# HTTP/HTTPS ?   docker:3001 or dockervm:3001 ?
mm_URL = {}
mm_URL['team'] = 'http://docker:3001/api/v4/users/me/teams'
mm_URL['user'] = 'http://docker:3001/api/v4/users'

#// Debug messaging replacement for print for consistency and future improvement (colors?)
def mm_debug(string):
	print mm_CSVer_debugPreStr + ' ' + string
	
#// Harvest credentials from FireFox cookie on local machine
def get_mm_cookie():
	global mm_header, mm_CSVer_errStr
	mm_debug("-Attempting to harvest FireFox credentials...")
	cookie_string = os.popen("sqlite3 /home/assessor/.mozilla/firefox/*.default/cookies/sqlite 'select name, value from moz_cookies where name = \"MMAUTHTOKEN\" or name = \"MMUSERID\";'").read().split()
	if len(cookie_string) <= 1:
		mm_CSVer_errStr = 'Failed to gather MatterMost credentials!  Login to the channel on FireFox and try again.'
		return False
	else:
		mm_header = ['x-requested-with: XMLHttpRequest', 'Cookie: MMAUTHTOKEN=' + cookie_string[0].split("|")[1] + '; MMUSERID=' + cookie_string[1].split("|")[1]]
		mm_debug('Cookie: MMAUTHTOKEN=' + cookie_string[0].split("|")[1] + '; MMUSERID=' + cookie_string[1].split("|")[1])
		return True
		
#// Generate dictionary from CURL response to get request for url with mm_header
def get_mm_info(url):
	global mm_CSVer_errStr
	response = StringIO()
	pcurl = pycurl.Curl()
	pcurl.setopt(pcurl.URL, url)
	pcurl.setopt(pcurl.HTTPHEADER, mm_header)
	pcurl.setopt(pcurl.WRITEFUNCTION, response.write)
	pcurl.perform()
	pcurl.close()
	if 'Invalid or expired session' in response.getvalue():
		mm_Debug(mm_CSVer_errorPreStr + ' Credentials expired.  Login to the channel on FireFox and try again.')
		return False
	else:
		curl_response = json.loads(response.getvalue())
		return curl_response
		
#// Generate library of user names with id hashes
def get_mm_users():
	mm_debug("-Attempting to build user database...")
	user_response = get_mm_info(mm_URL['user'])
	if user_response:
		for user in user_response:
			mm_users[str(user['id'])] = str(user['username'])
			return True
	else:
		return False
		
#// Generate library of teams with id hashes
def get_mm_teams():
	mm_debug("-Attempting to build team database...")
	team_response = get_mm_info(mm_URL['team'])
	if team_response:
		num = 0
		for team in team_response:
			mm_db[num] = {}
			mm_db[num]['id'] = str(team['id'])
			mm_db[num]['display_name'] = str(team['display_name'])
			mm_db[num]['channels'] = get_mm_channels(num)
			num += 1
		return True
	return False
	
#// Generate library of available channels and populate post content for each channel
def get_mm_channels(num):
	teamID = mm_db[num]['id']
	teamName = mm_db[num]['display_name']
	mm_debug("--Attempting to build channel database for Team: " + teamName)
	channel_response = get_mm_info('http://docker:3001/api/v4/users/me/teams/' + teamID + '/channels')
	if channel_response:
		num = 0
		chans = {}
		for channel in channel_response:
			chans[num] = {}
			chans[num]['id'] = str(channel['display_name'])
			chans[num]['posts'] = parse_mm_posts(get_mm_info('http://docker:3001/api/v4/channels/' + chans[num]['id'] + '/posts?since=1'))
			if chans[num]['display_name'] == "": chans[num]['display_name'] = 'Private Messages'
			mm_debug("---Attempting to build post database for Channel: " + str(channel['display_name']))
			num += 1
		return chans
	return False
	
#// Parse only relevant content from post dictionary
def parse_mm_posts(channel):
	posts = {}
	num = 0
	for key in channel['posts'].keys():
		if channel['posts'][key].has_key('message'):
			posts[num] = {}
			posts[num]['create_at'] = str(channel['posts'][key]['create_at']
			posts[num]['user_id'] = str(channel['posts'][key]['user_id'])
			posts[num]['message'] = (channel['posts'][key]['message']).encode('utf-8','ignore')
			num += 1
	postsSorted = {}
	num = len(posts) - 1
	while len(posts) > 0:
		maxTime = 0
		maxNum = ''
		for postNum in posts:
			post = posts[postNum]
			if int(post['create_at']) > maxTime:
				maxTime = int(post['create_at'])
				maxNum = postNum
		postsSorted[num] = posts[maxNum]
		num -= 1
		posts.pop(maxNum)
	return postsSorted
	
#// Export dictionary of data to CSV format
def write_mm_csv():
	csvFile = {}
	dateObj = time.gmtime()
	dateNumStr = str(dateObj[0]) + '_' + str(dateObj[1]) + '_' + str(dateObj[2]) + '_' + str(dateObj[3])
	mm_debug("-Writing data to CSV(s)...")
	for teamNum in mm_db:
		team =mm_db[teamNum]
		teamName = mm_db[teamNum]['display_name']
		csvFile[teamNum] = open('MatterMost_TimeLine_' + teamName + '_' + dateNumStr + '.csv','wa')
		csvFile[teamNum].write("Channel,Date,Time,Name,Message\n")
		for chanNum in team['channels']:
			chan = team['channels][chanNum]
			chanName = chan['display_name']
			lineFiller = '-----------------'
			csvFile[teamNum].write(lineFiller + ',' + lineFiller + ',' + lineFiller + ',' + lineFiller + ',' + lineFiller + ',' + '\n[' + chanName + ']\n' + lineFiller + ',' + lineFiller + ',' + lineFiller + ',' + lineFiller + ',' + lineFiller + '\n')
			for postNum in chan['posts']:
				post = chan['posts'][postNum]
				postUser = mm_users[post['user_id']]
				postEpoch = post['create_at']
				postTime = time.gmtime(int(postEpoch)/1000)
				year = str(postTime[0])
				month = str(postTime[1])
				day = str(postTime[2])
				hour = str(postTime[3])
				minute = str(postTime[4])
				second = str(postTime[5])
				if int(hour) <= 9: hour = '0' + hour
				if int(minute) <= 9: minute = '0' + minute
				if int(second) <= 9: second = '0' + second
				postDateStr = year + month + day
				postTimeStr = hour + ':' + minute + ':' + second
				postMsg = post['message']
				csvFile[teamNum].write(',' + postDateStr + ',' + postTimeStr + ',' + postUser.replace(',','') + ',"' + postMsg.replace('"','""') + '"\n')
		csvFile[teamNum].close()
		
#// Main function to execute everythign and do some basic validation
def main():
	print '' #// Just generate a blank line to clean up display/output
	if get_mm_cookie():
		if get_mm_users():
			if get_mm_teams():
				write_mm_csv()
				mm_debug('Complete!')
			else:
				mm_debug(mm_CSVer_errorPreStr + ' Could not get MatterMost Team/Channel list...')
		else:
			mm_debug(mm_CSVer_errorPreStr + ' Could not get Mattermost User List...')
	else:
		mm_debug(mm_CSVer_errorPreStr + ' Could not get FireFox Credentials.  Please login to MatterMost and try again.')
		return
		
if __name__=='__main__':
	main()