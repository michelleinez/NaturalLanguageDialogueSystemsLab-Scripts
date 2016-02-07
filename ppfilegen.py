import sys, os, re, string
from docxConv.docx import opendocx, getdocumenttext

'''
def createTxtFiles():
	creates a directory of txt files from all docx files in working directory
'''
def createTxtFiles():
	#this gives the path of the working directory of this script
	filesList = os.listdir(os.path.dirname(os.path.abspath(__file__)))

	for dirFile in filesList:
		search = re.search('\.docx', dirFile)
		#if the file is a docx then create a txtfile for it in txtFiles directory
		if search:
			try:
				docx = opendocx("./"+dirFile)
				if not os.path.exists("./txtFiles/"):
					os.makedirs("./txtFiles/")
				tempfile = open("./txtfiles/"+dirFile[:-5]+"temp.txt", 'w')
			except:
			    print(
			        "Failed to create .txt files."
			    )
			    exit()

			# Fetch all the text out of the document we just created
			paratextlist = getdocumenttext(docx)

			# Make explicit unicode version
			newparatextlist = []
			for paratext in paratextlist:
			    newparatextlist.append(paratext.encode("utf-8"))

			# Print out text of document
			tempfile.write(''.join(newparatextlist))
			tempfile.close()

			# Open the file again to ensure proper spacing and ordering of speakers
			tempfile = open("./txtfiles/"+dirFile[:-5]+"temp.txt", 'r')
			lines = ''
			t=2
			# t%2==0 means A is talking, t%2!=0 means B is talking

			for line in tempfile:
				#split the lines by dialogue turn...
				line=re.split('([A-B]\d:)', line.strip())
				for l in line:
					if (t%2==0):
						rgx = re.compile('A%s'%str((t/2)))
					else:
						rgx = re.compile('B%s'%str((t/2)))

					if re.match(rgx,l):
						t+=1
						lines+=(l)
					elif l and not re.match('\w',l):
						lines+=(l+"\n\n")
					else:
						continue
			tempfile.close()

			newfile = open("./txtfiles/"+dirFile[:-5]+".txt", 'w')
			newfile.write(lines)
			newfile.close()



			os.remove("./txtfiles/"+dirFile[:-5]+"temp.txt")
	return

'''
def createPhaseFile(speaker, fname):
	opens and adds header contents to phase file to prepare for appending of gestures
'''
def createPhaseFile(speaker, fname):
	phaseFile = open(os.path.dirname(os.path.abspath(__file__))+"/phasePhraseFiles/"+(fname[:-4])+"_"+speaker+".phase", 'w')
	hcontents = "#correct order: prep, stroke, hold, retract.\n\n"
	hcontents += "#prep: to move from default gesture to begin position\n"
	hcontents += "#stroke: perform the gesture\n"
	hcontents += "#hold: stay at gesture end position\n"
	hcontents += "#retract: move back tp default gesture position\n\n"
	hcontents += "#must have stroke\n"
	phaseFile.write(hcontents)
	phaseFile.close()
	return

'''
def createPhraseFile(speaker, fname):
	opens and adds header to phrase file to prepare for appending of gestures
'''
def createPhraseFile(speaker, fname):
	phraseFile = open(os.path.dirname(os.path.abspath(__file__))+"/phasePhraseFiles/"+(fname[:-4])+"_"+speaker+".phrase", 'w')
	phraseFile.write("TimeOffset 0.0\n"+"start end lexeme handedness path handshape hand-height-1 hand-body-dist-1 hand-radial-orient-1 elbow-inclination-1 hand-height-2 hand-body-dist-2 hand-radial-orient-2 elbow-inclination-2 2H-distance-start 2H-distance shoulders dss_lex_affil dse_lex_affil lex_affil dss_cooc dse_cooc cooc")
	phraseFile.close()
	return

'''
def writeToPhaseFile(speaker, filename, contents):
	appends lines of gestures to phase file
'''
def writeToPhaseFile(speaker, filename, contents):
	phaseFile = open(os.path.dirname(os.path.abspath(__file__))+"/phasePhraseFiles/"+filename[:-4]+"_"+speaker+".phase", 'a')
	phaseFile.write(contents)
	phaseFile.close()
	return

'''
def writeToPhraseFile(speaker, filename, contents):
	appends lines of gestures to phrase file
'''
def writeToPhraseFile(speaker, filename, contents):
	phraseFile = open(os.path.dirname(os.path.abspath(__file__))+"/phasePhraseFiles/"+filename[:-4]+"_"+speaker+".phrase", 'a')
	phraseFile.write(contents)
	phraseFile.close()
	return

'''
def grabGesturesFromLine(speaker, txtFileName, lineNumber, line):
	takes lines in as string and outputs a list of lists in the form of
	[[gesture, hands],[gesture,hands]...]
'''
def grabGesturesFromLine(speaker, txtFileName, lineNumber, line):
	writeToPhaseFile(speaker, txtFileName, "\n#"+speaker+str(lineNumber))
	gesturesAdptd=[]
	gesturesNonAdptd=[]
	gesturechunks=re.findall('\(\[.*?\]\)',line)
	if gesturechunks:
		for gesturechunk in gesturechunks:
			gestureNum=0
			g1=None
			g2=None
			gesturelist = re.findall('\].+?\[', gesturechunk)
			for gesture in gesturelist:
				gestureNum+=1
				gesture = re.sub('[\[\]/ ]', '', gesture)
				gesture=(re.split('[,_](2H|RH|LH)',gesture))
				if not re.match('(2H|RH|LH)',gesture[1]) or not re.match('',gesture[0]):
					print "error: something wrong with this annotation..."
					quit()
				#get rid of all empty list elements after the split
				while True:
					try:
						gesture.remove("")
					except ValueError:
						break
				if gestureNum==1:
					g1=gesture
					print "g1="+str(g1)
				elif gestureNum==2:
					g2=gesture
					print "g2="+str(g2)
				else:
					print 'error with gesture: '+str(gesturechunk)
					quit()
			gesturesAdptd.append(g1)
			if g2:
				gesturesNonAdptd.append(g2)
			else:
				gesturesNonAdptd.append(g1)
	else:
		print 'error: there are no gestures on this line.'
		quit()
	return [gesturesAdptd, gesturesNonAdptd]
'''
def grabTimestampsFromLine(line):
	takes lines in as string and outputs an adapted and nonadapted list of each gestures' timestamps in the form of
	[[startTime,duration],[startTime,secondDuration]...]
'''
def grabTimestampsFromLine(line):
	timestampsNotAdptd=[]
	timestampsAdptd=[]
	#find all gesture chunks in the line
	gesturechunks=re.findall('\(\[.*?\]\)',line)
	#for each gesture chunk
	for gesturechunk in gesturechunks:
		#create a list of adapted timestamps and nonadapted timestamps
		ts=re.findall('\[.+?\]', gesturechunk)
		identity = string.maketrans("", "")
		ts = [s.translate(identity, "[\[\]]") for s in ts]
		ts2=[]
		print "pre split ts="+str(ts)
		print "pre split ts2="+str(ts2)
		#if there are three timestamps in the chunk (adapted/nonadapted annotation)
		if len(ts)==3:
			#then store the second timestamp start and duration in ts2[]
			#for the second gesture
			ts2.append(ts[0])
			ts2.append(ts[2])
			ts=ts[:-1]
		print "post split ts="+str(ts)
		print "post split ts2="+str(ts2)
		#if a second timestamp exists
		if(ts2):
			#append the first timestamp (ts) to the adapted list
			timestampsAdptd.append(ts)
			#and append the second timestamp (ts2) to the nonadapted list
			timestampsNotAdptd.append(ts2)
		else:
			#otherwise append the first timestamp to both the adapted and nonadapted list
			timestampsAdptd.append(ts)
			timestampsNotAdptd.append(ts)
	print "timestampsAdptd="+str(timestampsAdptd)
	print "timestampsNotAdptd="+str(timestampsNotAdptd)
	return [timestampsAdptd,timestampsNotAdptd]

'''
def processTxtFiles():
	opens text files and harvests/sorts gestures
	(invokes the above functions)
'''
def processTxtFiles():

	#create directory for phase/phrase files
	if not os.path.exists("./phasePhraseFiles/"):
		os.makedirs("./phasePhraseFiles/")

	#get each of the .txt filenames and open them consecutively...
	txtFileList = os.listdir(os.path.dirname(os.path.abspath(__file__))+"/txtFiles")
	#there are speakers A and B
	speakers=["A", "B"]
	gestureNumber=0
	#so do the following twice--once for each speaker
	for speaker in speakers:
		#create a gesture list gestureList[]
		gestureList=[]
		timestampList=[]
		#and then for every textfile name
		for txtFileName in txtFileList:

			#create a phase file and a phrase file for the iterated speaker
			#this file gets named such as the following example: 'bonfire1_A.phase'
			createPhraseFile(speaker, txtFileName)
			createPhaseFile(speaker, txtFileName)

			#initiate the lineNumber counter
			lineNumber = 0
			#start a list of timestamps called rawTimes[]
			rawTimes = []
			#open txtfile for reading
			txtFileContents = open(os.path.dirname(os.path.abspath(__file__))+"/txtFiles/"+txtFileName, 'r')
			#set speakingRGX as the regular expression for A1:, B5: etc.
			speakingRGX = r'{0}\d+:'.format(speaker)
			#then for every line in the text file
			for line in txtFileContents:
				#count which gesture chunk we are on...
				gestureChunkNum=0
				#if this line of the file contains speech by the speaker we are iterating over
				#then record the speaker and the turn number of the dialogue turn
				if re.match(speakingRGX, line):
					dturn=(re.search(speakingRGX, line)).group()
					turnNum=dturn[1:-1]
					lineNumber += 1
					print speaker+" line "+str(lineNumber)
					print "=============================\n"+line+"=============================\n"
				#but if the line is not spoken by the speaker we are iterating over
				#then continue on to the next line and don't do any more parsing for this line
				else:
					continue
				gestures=grabGesturesFromLine(speaker, txtFileName, lineNumber, line)
				gesturesAdptd=gestures[0]
				gesturesNonAdptd=gestures[1]
				print "gesturesAdptd = "+str(gesturesAdptd)
				print "gesturesNonAdptd = "+str(gesturesNonAdptd)
				timestamps=grabTimestampsFromLine(line)
				timestampsAdptd=timestamps[0]
				timestampsNotAdptd=timestamps[1]


				Adapted=zip(gesturesAdptd,timestampsAdptd)
				NonAdapted=zip(gesturesNonAdptd,timestampsNotAdptd)
				print "Adapted="+str(Adapted)
				print "NonAdapted="+str(NonAdapted)
				phaseContents=str(Adapted)
				for gesture in Adapted:
					gestureNumber+=1
					gestureStartTime=float(gesture[1][0][:-1])
					print "gestureStartTime="+str(gestureStartTime)
					gestureDuration=float(gesture[1][1])
					print "gestureDuration="+str(gestureDuration)

					#Phase file writing...
					writeToPhaseFile(speaker, txtFileName, "\n#"+str(gestureNumber)+" "+str(gesture[0][0]))
					writeToPhaseFile(speaker, txtFileName, "\n"+str(gestureStartTime-0.02)+" "+str(gestureStartTime)+" prep")
					writeToPhaseFile(speaker, txtFileName, "\n"+str(gestureStartTime)+" "+str(gestureStartTime+gestureDuration)+" stroke")
					writeToPhaseFile(speaker, txtFileName, "\n"+str(gestureStartTime+gestureDuration)+" "+str(gestureStartTime+gestureDuration+0.2)+" retract\n")

					writeToPhraseFile(speaker, txtFileName, "\n\n"+str(gestureStartTime)+" "+str(gestureStartTime+gestureDuration))
					writeToPhraseFile(speaker, txtFileName, " "+str(gesture[0][0]+" "+str(gesture[0][1])))
			txtFileContents.close()
	return

createTxtFiles()
processTxtFiles()