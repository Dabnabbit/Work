#!/bin/python
#//- This script's purpose is to automate the distribution and management of scripts across multiple
#//-	nodes/servers/sensors. It allows you to copy/distribute, run, or kill scripts from a single
#//-	command line execution across numerous systems.

import os, sys, re, json, time

logPath = '/home/assessor/'

startTime = 0.0

def main():
	args = sys.argv
	args.pop(0) #//- Remove script name from args
	#print "ARGS:",args
	if len(args) >= 5:
		fileName = str(args.pop(0))
		userName = str(args.pop(0))
		fileDir = str(args.pop(0))
		optExec = str(args.pop(0))
		if fileDir[-1] != '/':
			fileDir += '/'
		if not os.path.isfile(fileName):
			print "[X] Read Error:\n Script: " + fileName + " does not exist..."
			return False
		for host in args:
			if "copy" in optExec:
				cmd = "scp "+fileName+" "+userName+"@"+host+":"+fileDir+"."
				os.popen(cmd)
				print host+': File '+fileDir+fileName+' Copied.'
			if "status" in optExec:
				cmd = 'ssh ' + userName + '@' + host + ' ps -ef | grep '+fileDir+fileName+' | grep -v grep'
				resp = str(os.popen(cmd).read()).strip().split('\n')				
				pids = ''
				statusStr = {}
				if len(resp) and resp[0] != '':
					for line in resp:
						data = line.split()
						if fileDir+fileName in data[-1]:
							pids += ' '+data[1].strip()
							statusStr[data[1].strip()] = data[4].strip()
				#pids = pids.strip()
				if pids != '':
					#os.popen('ssh ' + userName + '@' + host + ' kill'+pids)
					print host+': Process '+fileDir+fileName+' Found:'
					for pID in pids.split():
						print 'PID: '+pID+'\tStartTime: '+statusStr[str(pID)]
				else:
					print host+': Process '+fileDir+fileName+' not found.'
			if "kill" in optExec:
				cmd = 'ssh ' + userName + '@' + host + ' ps -ef | grep '+fileDir+fileName+' | grep -v grep'
				resp = str(os.popen(cmd).read()).strip().split('\n')
				pids = ''
				if len(resp) and resp[0] != '':
					for line in resp:
						data = line.split()
						if fileDir+fileName in data[-1]:
							pids += ' '+data[1]
				if pids != '':
					os.popen('ssh ' + userName + '@' + host + ' kill'+pids)
					print host+': Process '+fileDir+fileName+' -- PID:'+pids+' Killed.'
				else:
					print host+': Process '+fileDir+fileName+' not found.'
			if "run" in optExec:
				sshCmd = 'nohup /bin/python '+fileDir+fileName+' &'
				cmd = 'ssh ' + userName + '@' + host + ' '+sshCmd
				os.popen(cmd)
				cmd = 'ssh ' + userName + '@' + host + ' ps -ef | grep '+fileDir+fileName+' | grep -v grep'
				resp = str(os.popen(cmd).read()).strip().split('\n')
				if len(resp) and resp[0] != '':
					print host+': Process '+fileDir+fileName+' -- PID: '+resp[0].split()[1]+' Started'
	else:
		print "[X] Syntax Error:\n script.py fileName userName directory copy/status/kill/run host1 host2 host3...\n"			
		return False
		
if __name__ == "__main__":
	startTime = time.gmtime()
	main()

