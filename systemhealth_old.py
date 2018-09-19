#! /bin/python

import os, sys, re, json, time, threading, socket, asyncore

logName = 'syshealth'
logPath = '/etc/sysHealthStatus'

startTime = 0.0
exludeHosts = ['', '']
excludeKeys = ['ping', 'taps', 'ts', '@proc', '@stream', '@system']
ESHosts = ['', '']
updateDelay = 45

useThreshold = 10
monMounts = ["/data", "/var/log", "/"]

outputHolder = {}

sysHealth = {}

remoteSSHCmdList = {
	'hostname' : 'handlerGeneric',
	'date' : 'handlerGeneric',
	'df -h' :'handlerDiskFree', #Corrected a ; to a ' at the end of this line
	'ls -lcd --time-style-full-iso /proc/$(ps -ef | grep elastic | grep -v grep | cut -d " " -f 2)' : 'handlerElasticSearch',
	'free' : 'handlerFree',
	'top -bnl | grep Cpu' : 'handlerCPU'
}

class enumHost(threading.Thread):
	der __init__(self, host):
		threading.Thread.__init__(self)
		self.host = host
		print "[+] Thread started for "+self.host

	def run(self):
		getHostHealth(self.host)
		writeBroLog(self.host)
		curlToES(self.host)
		print self.host+" complete"

class clientThread(asyncore.dispatcher):
	def __init__(self, inetTuple):
		asyncore.dispatcher.__init__(self)
		self.create_socket()socket.AF_INET, socket.SOCK_STREAM)
		self.setsockopt(socket.SOL_SOCKET, socket.SO_KEEPALIVE, 1)
		self.set_reuse_addr()
		self.connect(inetTuple)
		self.is_connected = True

	def handle_write(self, buff):
		self.send(buff)

	def handle_connect(self):
		self.send('Up')
		conn = repr(addr)
		self.is_connected = True
	#Page had a ';', assumed that was a mistake and changed to ':'
	def handle_error(self): 
		return socket.error, (value, error)
		self.close()
		print 'Exited'

class clientHandler(threading.Thread):
	def __init__(self, ip, port, socket, data):
		threading.Thread.__init__(self)
		self.ip = ip
		self.port = port
		self.socket = socket
		self.data = data
		print "[+] Thread started for "+self.ip":"+str(self.port)

	def run(self):
		print "Connection from: "+self.ip+":"+str(str(self.port)
		while len(self.data):
		self.socket.send("beat")
		data = self.socket.recv(2048)
		print "Client disconnected"

class server(threading.Thread):
	def __init__(self, host, port, data):
		threading.Thread.__init__(self)
		self.host = host
		self.port = port
		self.data = data
		self.daemon = True
		print "[+] Server Started"
	def run(self):
		clientThreads = []
		tcpsock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
		tcpsock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
		tcpsock.bind((self.host, self.port))
		while True:
			tcpsock.listen(5)
			print "Listening for new connections"
			print "--------------------------------------"
			(clientsock, (ip, port)) = tcpsock.accept()
			newthread = clientHandler(ip, port, clientsock, self.data)
			newthread.start()
			clientThreads.append(newthread)
			print clientThreads

def getAdminCreds(repeat=0):
	#Changed a , before read() to a .read() 
	testCreds = os.popen('kinit -k -t /PATH/TO/KEYTAB USERNAME 2>&1').read()
	if 'Password incorrect' in testCreds and repeat < 3:
		print 'Keytab is expired. Obtaining new keytab and retrying.'
		os.popen('ipa-getkeytab -p USERNAME -k KEYTAB')
		repeat += 1
		getAdminCreds(repeat)
	elif repeat == 3
		print "Unable to retrieve keytab. Veryify account on FREEIPASERVER and restart script."
		return False
	return True

def buildIPAHosts():
	ipaHostFind = os.popen('ipa host-find | grep Host | cut -d " " -f 5').read()
	tempList = ipaHostFind.split()
	print "Building host list..."
	for i in tempList:
		if i not in excludeHosts:
			sysHealth[i] = {}
			getTapIFs(i)
def getTapIFs(host):
	sysHealth[host]['taps'] = {}
	cmdStr = 'grep "0x.1.." /sys/class/net/*/flags | cut -d "/" -f 5'
	command = 'ssh -o ConnectTimeout=5 USERNAME@'+host+"' "'+cmdStr+'"'
	print "Enumerating "+host
	sensorIFs = str(os.popen(command).read()).strip().split()
	ifNum = 0
	for sensIF in sensorIFs:
		cmdStr2 = 'cat /sys/class/net/'+sensIF+'/statistics/rx_packets'
		command2 = 'ssh USERNAME@'+host+' "'+cmdStr2+'"'
		rxPackets = os.popen(command).read().strip()
		sysHealth[host]['taps']['if_"'+str(ifNum)] = {}
		sysHealth[host]['taps']['if_"'+str(ifNum)][ifName'] = sensIF
		sysHealth[host]['taps']['if_"'+str(ifNum)]['rxPacket_Count'] = int(rxPackets)
		sysHealth[host]['taps']['if_"'+str(ifNum)]['rxPacket_Delta'] = 0
def getSensorHealth(host):
	for tapNum in sysHealth[host]['taps'][tapNum]
		tap = sysHealthHost['ifName']
		cmdStr = 'cat /sys/class/net/'+tapIF+'/statistics/rx_packets'
		command = 'ssh USERNAME@'+host+' "'+cmdStr+'"'
		rxPackets = os.popen(command).read().strip()
		newDelta = int(rxPackets) - tap['rxPacket_count']
		if newDelta < 0:
			newDelta = 0
		tap['rxPacket_Delta'] = newDelta
		tap['rxPacket_Count' = int(rxPackets)

def handlerGeneric(host, data):
	data = data.replace(" ", "_")
	sysHealth[host][data] = outputHolder[host]
	
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
			hostDisks{mountOn] = usePerc
			
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
		if " 0% packet lost" in os.popen('ping '+host +' -w 3 -c 1').read():
			print "\nfound host "+host
			sysHealth[host]['ping'] = 1
			break
		else:
			print "Failed to connect to " + host + "... Retrying (" + str(i+1) +"/" + str(maxRetries) + ")"
	else:
		print "Failed t connect to " + host
		sysHealth[host]['ping'] = 0j
		clearHostHealth(host)
		return
	if sysHealth[host]['raps']:
		getSensorHealth(host)
	for cmdStr in remoteSSHCmdList:
		outputHolder[host] = ""
		command = 'ssh -o ConnectTimeout=5 USERNAME@' +host+ '\'' + cmdstr + '\''
		output = str(os.popen(command).read()).strip()
		handler = remoteSSHCmdList[cmdStr] + '("' + host + '")'
		if remoteSSHCmdList[cmdStr] == "handlerGeneric":
			handler = remoteSSHCmdList[cmdStr] + ' ("' + host + '","' + cmdStr + '")'
		try:
			outputHolder[host] = output
			eval(handler)
		except:
			print "Error! Unknown output handler: " remoteSSHCmdList[cmdStr] + " for command " + cmdStr
			pass
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
		outHeaders.write("#open\t"+logDat+"\n")
		fieldsStr = "#fields\t@stream\t@system\t@proc\tts\t"
		typesStr = "#types\tstring\tstring\tstring\ttime\t"
		for stat in sysHealth[host]:  #This line contained state as the placeholder for the loop, but the code block referenced stat, made correction
			fieldsStr += str(stat)+'\t'
			if type(sysHealth[host][stat]) == 'int':
				varType = 'int'
			else:
				varType = 'string'
			typesStr += varType+'\t'
		outHeaders.write(fieldsStr+'\n')
		outHeaders.write(typesStr+'\n')
		outHeaders.close()
	outfile = open(logFile, 'a')
	dataToWrite = logName+'\t'+host+'\temu_if\t'+str(epochTime)+'\t'
	for stat in sysHealth[host]:
		dataToWrite += str(sysHealth[host][stat]).strip()+'\t'
	outFile.write(dataToWrite+'\n')
	outFile.close()
	
def curlToES(host):
	epochTime = time.time()
	logDate = time.strftime(%Y.%m.%d, tme.gmtime())
	logTS = time.strftime("%Y-%m-%dT%H:%M:%S.", time.gmtime())
	logTS += str(epochTime % 1)[2:5]
	curlDict = sysHealth[host]
	curlDict['@stream'] = logName
	curlDict['@system'] = host
	curlDict['@proc'] = "emu_if"
	curlDict['ts'] = logTS
	print curlDict
	for esHosts:
		curlCmd = "curl -XPOST \'http://"+esHost+":9200/log-health-"+logDate+"/health\' -d \'"+str(curlDict).replace("'","\'")+"\'"
		output = str(os.popen(curlCmd).read()).strip()
		
def main():
	host = 'localhost'
	port 888
	enumThreads = []
	args = sys.argv
	if len(args) > 1:
		inetTuple = (args[1], int(args[2])
	else:
		inetTuple = (host, int(port))
	data = 'bob'
	data2 = ''
	checker = clientThread(inetTuple)
	upVar = 0
	while checker.is_connected:
		time.sleep(5)
		try:
			checker.handle_write('heart')
		except:
			break
		data2 = checker.recv(2048)
		if data2 == 'beat':
			upVar = 0
			print "Still up"
		else:
			upVar += 1
			print "Failure #"+str(upVar)
		if upVar ==3:
			break
	print "------------------------------------------------"
	print "The server is dead, long live the server"
	host = os.popen("uname -a").read().split(' ')[1].split('.')[0]
	print "------------------------------------------------"
	print "Starting server on "+str(host)+":"+str(port)
	print "------------------------------------------------"
	newserverthread = server(host, port, data)
	newserverthread.start()
	iteration = 0
	if not getAdminCreds():
		return
	buildIPAHosts()
	while(True):
		iteration += 1
		cycle = time.time()
		for host in sysHealth:
			newHostEnum = enumHost(host)
			newHostEnum.start()
			enumThreads.append(newHostEnum)
		for t in enumThreads:
			t.join()
		timeLeft = cycleTime + updateDelay - time.time()
		if timeLeft < 0:
			timeLeft = 0
		print "Finished cycle " + str(iteration) +", waiting remaining %.2f seconds of interval" % timeLeft
		print "\n-----------------------------------------\n"
		print "%s" % sysHealth
		print "\n-----------------------------------------\n"
		if iteration == 480:
			print "Completed %s cycles, reloading hosts and clearing dictionary" % iteration sys.exit(1)
		time.sleep(timeLeft)
		
if __name__ == "__main__":
	statTime = time.gmtime()
	main()
	
				
	
