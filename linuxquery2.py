#!/bin/python

import os, sys, re, json, time

logName = 'query'
logPath = '/etc/sysHealthStatus/'

startTime = 0.0
excludeHosts = ['', '']
excludeKeys = ['ping', 'taps', 'timestamp', '@proc', '@stream', '@system']
#ESHosts = ['elasticnodes-0.dmss', 'elasticnodes-1.dmss', 'elasticnodes-2.dmss', 'elasticnodes-3.dmss', 'elasticnodes-4.dmss', 'elasticnodes-5.dmss']
ESHosts = ['elasticnodes-0.dmss'] #//- Will incorporate load balancing here later
updateDelay = 10
userName = "USERNAME"

#//- Enable indexing per command vs per host
indexPerCmd = True


#// HDD values, mostly for dashboard stuff
useThreshold = 10
monMounts = ["/data", "/var/log", "/"]

outputHolder = {}

sysHealth = {}

remoteSSHCmdList = {}
remoteSSHCmdListOld = {
	'hostname' : 'handlerGeneric',
	'date' : 'handlerGeneric',
	'df -h' :'handlerDiskFree', #Corrected a ; to a ' at the end of this line
	'ls -lcd --time-style-full-iso /proc/$(ps -ef | grep elastic | grep -v grep | cut -d " " -f 2)' : 'handlerElasticSearch',
	'free' : 'handlerFree',
	'top -bnl | grep Cpu' : 'handlerCPU'
}

grok_Types = {
	'WORD': r'\w+',
	'NUMBER': r'\d+'
}

def buildSSHCmdList(cmdFileName):
	print "[+] Building SSH Command List:"
	try:
		cmdFile = open(cmdFileName, 'r')
		cmdLines = cmdFile.readlines()
		for line in cmdLines:
			lineSplit = line.split(":",1)
			print " - \'" + lineSplit[0] + "\' : " + lineSplit[1].strip()
			remoteSSHCmdList[lineSplit[0]] = lineSplit[1].strip()
		#print remoteSSHCmdList
	except:
		print "[X] Error parsing " + cmdFileName
		return False

def buildIPList(ipFileName):
	print "[+] Building IP List:"
	try:
		ipFile = open(ipFileName, 'r')
		ipLines = ipFile.readlines()
		for ip in ipLines:
			print " - " + ip.strip()
			sysHealth[ip.strip()] = {}
	except:
		print "[X] Error parsing " + cmdFileName
		return False

def formatCmdName(cmd):
	return cmd.replace(" ", "_").lower()

#//- GROK functionality still expiremental at best, pretty complicated but works in theory... just not 100% sure it will work for all outputs
def grokMakePattern(pat):
	return re.sub(r'%{(\w+):(\w+)}',
		lambda m: "(?P<" + m.group(2) + ">" + grok_Types[m.group(1)] + ")", pat)

#//- Need to finish this bit eventually with valid field naming etc
def handlerGROK(host, data):
	print "Pattern: \'" + data + "\'"
	pattern = grokMakePattern(data)
	print re.search(pattern, outputHolder[host]).groupdict()


#//- Regular default handler for data
def handlerGeneric(host, data):
	data = formatCmdName(data)
	sysHealth[host][data] = outputHolder[host]
	#sysHealth[host]['size'] += len(outputHolder[host])
	
def handlerFree(host):
	lines = outputHolder[host].split('\n')
	mem = lines[1].split()
	swap = lines[2].split()
	sysHealth[host]['free_mem'] = mem[3]
	sysHealth[host]['free_swap'] = swap[3]
	sysHealth[host]['mem_free_perc'] = round(float(float(mem[1]) / float(mem[1]) * 100), 2)
	
def handlerCPU(host):
	line = outputHolder[host].split(',')
	sysHealth[host]['cpu_idle'] = round(float(line[3].split()[0]), 2)
	
def handlerDiskFree(host):
	sysHealth[host]['disk'] = {}
	hostDisks = sysHealth[host]['disk']
	lines = outputHolder[host].split('\n')
	lines.pop(0)
	for line in lines:
		data = line.split()
		usePerc = int(data[4][:-1])
		mountOn = data[5]
		if usePerc > useThreshold and mountOn in monMounts:
			hostDisks[mountOn] = usePerc
			
def handlerProcess(host):
	sysHealth[host]['process'] = {}
	hostProcs = sysHealth[host]['process']
	lines = outputHolder[host].split('\n')
	lines.pop(0)
	i = 0
	for line in lines:
		hostProcs[str(i)] = line
		i += 1
			
def handlerNetstat(host):
	sysHealth[host]['netstat'] = {}
	hostNetstat = sysHealth[host]['netstat']
	lines = outputHolder[host].split('\n')
	lines.pop(0)
	i = 0
	for line in lines:
		hostNetstat[str(i)] = line
		i += 1

def handlerElasticSearch(host):
	if re.search('/proc/$', outputHolder[host]):
		sysHealth[host]['ES'] = 'NotFound'
	elif re.search('/proc/[0-9]+', outputHolder[host]):
		sysHealth[host]['ES'] = 'Running'
		sysHealth[host]['ESUpSince'] = str(outputHolder[host].split(' ')[5] + 'T' + outputHolder[host].split(' ')[6].split('.')[0])

def clearHostHealth(host):
	for k in sysHealth[host].keys():
		if k in sysHealth[host] and k not in excludeKeys:
			del sysHealth[host][k]
	
def getHostHealth(host):
	global outputHolder
	maxRetries = 3
	for i in range(0, maxRetries):
		pingtest = os.popen('ping ' + host + ' -w 3 -c 1').read()
		if " 0% packet loss" in pingtest: #// So this response depends on the host ping is launched from...
			print "\n[+] Found host \'" + host + "\'"
			sysHealth[host]['ping'] = 1
			break
		else:
			print "\n[X] Failed to connect to \'" + host + "\'... Retrying (" + str(i+1) +"/" + str(maxRetries) + ")"
	else:
		print "[X] Failed to connect to " + host
		sysHealth[host]['ping'] = 0
		clearHostHealth(host)
		return
	#sysHealth[host]['size'] = 0
	for cmdStr in remoteSSHCmdList:
		outputHolder[host] = ""
		command = 'ssh -o ConnectTimeout=5 ' + userName + '@' + host + ' \'' + cmdStr + '\''
		print " - Attempting CMD: \'" + cmdStr + "\'"
		output = str(os.popen(command).read()).strip()
		handler = remoteSSHCmdList[cmdStr] + '("' + host + '")'
		if "GROK" in handler:
			pattern = remoteSSHCmdList[cmdStr][5:]
			handler = 'handlerGROK ("' + host + '","' + pattern + '")'
		elif remoteSSHCmdList[cmdStr] == "handlerGeneric":
			handler = remoteSSHCmdList[cmdStr] + ' ("' + host + '","' + cmdStr + '")'
		try:
			outputHolder[host] = output
			eval(handler)
		except:
			print "[X] Error! Unknown output handler: " + remoteSSHCmdList[cmdStr] + " for command " + cmdStr
			pass
		if indexPerCmd:
			curlToES(host,formatCmdName(cmdStr))
			clearHostHealth(host)


def writeBroLog(host):
	logFile = logPath+logName+'.log'
	epochTime = time.time()
	logDate = time.strftime("%Y-%m-%d-%H-%M-%S", time.gmtime())
	if not os.path.isfile(logFile):
		outHeaders = open(logFile, 'w')
		outHeaders.write("$separator \\x09\n")
		outHeaders.write("#set_separator\t,\n")
		outHeaders.write("#empty_field\t(empty)\n")
		outHeaders.write("#unset_field\t-\n")
		outHeaders.write("#path\t"+logName+"\n")
		outHeaders.write("#open\t"+logDate+"\n")
		fieldsStr = "#fields\t@stream\t@system\t@proc\tts\t"
		typesStr = "#types\tstring\tstring\tstring\ttime\t"
		for stat in sysHealth[host]:
			fieldsStr += str(stat)+'\t'
			if type(sysHealth[host][stat]) == 'int':
				varType = 'int'
			else:
				varType = 'string'
			typesStr += varType+'\t'
		outHeaders.write(fieldsStr+'\n')
		outHeaders.write(typesStr+'\n')
		outHeaders.close()
	outFile = open(logFile, 'a')
	dataToWrite = logName+'\t'+host+'\temu_if\t'+str(epochTime)+'\t'
	for stat in sysHealth[host]:
		dataToWrite += str(sysHealth[host][stat]).strip()+'\t'
	outFile.write(dataToWrite+'\n')
	outFile.close()
	
def curlToES(host,index=False):
	epochTime = time.time()
	logDate = time.strftime("%Y.%m.%d", time.gmtime())
	logTS = time.strftime("%Y-%m-%dT%H:%M:%S.", time.gmtime())
	logTS += str(epochTime % 1)[2:5]
	curlDict = sysHealth[host]
	curlDict['@stream'] = logName
	curlDict['@system'] = host
	curlDict['timestamp'] = logTS
	#print curlDict
	print len(curlDict)
	for esHost in ESHosts:
		indexName = logName + "-" + logDate
		if index:
			indexName = logName + "-" + index
		print " - Indexing \'" + host + "\' results to " + esHost + ":9200/" + indexName
		curlCmd = "curl -XPOST \'http://"+esHost+":9200/"+indexName+"/health\' -d \'"+str(curlDict).replace("'","\"")+"\'"
		#print curlCmd
		output = str(os.popen(curlCmd).read()).strip()
		if '"failed":0' in output:
			print " - Indexing Successful"
		else:
			print "[X] Error: Indexing failure"
			print output
		
def main():
	global userName
	args = sys.argv
	if len(args) > 3:
		#//- Syntax is script.py username cmdfile.txt ip1 ip2 ip3 ip...
		args.pop(0) #//- Remove script name from args
		userName = str(args.pop(0))
		cmdFileName = str(args.pop(0))
		if not os.path.isfile(cmdFileName):
			print "[X] Read Error:\n Cmd List: " + cmdFileName + " does not exist..."
			return False
		ipFileName = str(args.pop(0))
		if not os.path.isfile(ipFileName):
			print "[X] Read Error:\n IP File: " + ipFileName + " does not exist..."
			return False
		#for ip in args:
		#	if ip not in excludeHosts:
		#		sysHealth[ip] = {}
		if len(args):
			print "[X] Syntax Error:\n script.py username cmdListFile ipListFile\n"			
			return False
	else:
		print "[X] Syntax Error:\n script.py username cmdListFile ipListFile\n"
		return False
	buildSSHCmdList(cmdFileName)
	buildIPList(ipFileName)
	for host in sysHealth:
		getHostHealth(host)
		if not indexPerCmd:
			curlToES(host)
		#writeBroLog(host)
	#//- Core System Health Check loop, not needed for simple single-pass enumeration
	#iteration = 0
	#while(True):
	#	iteration += 1
	#	cycleTime = time.time()
	#	for host in sysHealth:
	#		getHostHealth(host)
	#		#writeBroLog(host)
	#		curlToES(host)
	#	timeLeft = cycleTime + updateDelay - time.time()
	#	if timeLeft < 0:
	#		timeLeft = 0
	#	print "Finished cycle " + str(iteration) + ", waiting remaining %.2f seconds of interval" % timeLeft
	#	print "\n-----------------------------------------\n"
	#	print "%s" % sysHealth
	#	print "\n-----------------------------------------\n"
	#	if iteration == 480:
	#		print "Completed %s cycles, reloading hosts and clearing dictionary" % iteration
	#	#	sys.exit(1)
	#	time.sleep(timeLeft)
		
if __name__ == "__main__":
	startTime = time.gmtime()
	main()

