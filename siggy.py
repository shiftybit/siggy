import sys
import os
import shutil
import win32api
import win32con
import win32com.client
import clr
import time
sys.path.append(os.getcwd())
import config
import notify
clr.AddReference("System.DirectoryServices")
import System.DirectoryServices
import hashlib
import pywintypes

def nprint(val): 
	pass
print1 = print
print2 = print
print3 = print
print4 = print

#Document Object -> https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.document_members.aspx

def SetupRegistry():
	try:
		HKCU = win32api.RegOpenCurrentUser()
		key = win32api.RegCreateKeyEx(HKCU,config.Settings['RegistryLocation'],win32con.KEY_READ | win32con.KEY_WRITE)
	except Exception as er:
		print1("Fatal, cannot access Registry %s" % config.Settings['RegistryLocation'])
		sys.exit(0)
	print1("First Time Run, Setting Up Registry")
	return key[0]
	
def GetDirectoryInformation():
	User = win32api.GetUserName()
	Searcher = System.DirectoryServices.DirectorySearcher()
	Searcher.Filter = "(&(objectCategory=User)(samAccountName=%s))" % User
	Searcher.PropertiesToLoad.Add("professionalcredentials")
	for Entry in config.ADProperties:
		print3("Adding %s" % Entry)
		Searcher.PropertiesToLoad.Add(Entry)
	
	try:
		SearchResult = Searcher.FindOne() #Type: System.DirectoryServices.SearchResult
		time.sleep(5)
		SearchResult = Searcher.FindOne() #Type: System.DirectoryServices.SearchResult
		time.sleep(5)
		SearchResult = Searcher.FindOne() #Type: System.DirectoryServices.SearchResult
	except Exception as er:
		print3("Could not connect to AD")
		return False
	PropertyCollection = SearchResult.Properties #Type: System.DirectoryServices.ResultPropertyCollection
	PropList = dict()
	for Entry in PropertyCollection:
		PropList[Entry.Key] = Entry.Value[0]
	#For any results not found, add the key to the PropList and set the value to None
	for Entry in config.ADProperties:
		if Entry not in PropList:
			PropList[Entry] = None
	return PropList

def GetRegistryInformation():	
	PropList = dict()
	HKCU = win32api.RegOpenCurrentUser()
	try:
		key = win32api.RegOpenKeyEx(HKCU,config.Settings['RegistryLocation'],0,win32con.KEY_READ)
	#FIX: Need to handle exceptions explicitly for win32api. 
	except Exception as er:
		key = SetupRegistry()
	for Entry in config.ADProperties:
		try:
			Value = win32api.RegQueryValueEx(key,Entry)
			PropList[Entry] = Value[0]
			if (Value[0] == ''):
				PropList[Entry] = None
		except Exception as er:
			print2("Reg: Could not find %s" % Entry)
			PropList[Entry] = None
	win32api.RegCloseKey(key)
	win32api.RegCloseKey(HKCU)
	return PropList
	
def IsDirectorySynchronized(RegPropList,ADPropList):
	#Fix. Need to handle Different TypeCasts between Registry and Directory
	ret = True
	for Entry in config.ADProperties:
		Reg = RegPropList[Entry]
		AD = ADPropList[Entry]
		print3("REG: %s \t AD: %s" % (Reg,AD))
		if(Reg == AD):
			print3("%s REG: %s \t AD: %s" % (Entry,Reg,AD))
		else:
			print1("%s REG: %s \t AD: %s" % (Entry,Reg,AD))
			ret = False
	return ret
	
def AddRegistryKey(PropertyName,Value):
	HKCU = win32api.RegOpenCurrentUser()
	key = win32api.RegOpenKeyEx(HKCU,config.Settings['RegistryLocation'],0,win32con.KEY_SET_VALUE)
	if(isinstance(Value,str) or Value is None): 
		win32api.RegSetValueEx(key,PropertyName,0,win32con.REG_SZ,Value)
		print3("Key Added -> Keytype %s. (%s,%s)" % (type(Value),PropertyName,Value))
	elif(isinstance(Value,System.DateTime)):
		print3("Date Time Type %s" % Value.ToString())
		win32api.RegSetValueEx(key,PropertyName,0,win32con.REG_SZ,Value.ToString())
	else:
		print1("unknown keytype %s. (%s,%s)" % (type(Value),PropertyName,Value))
	win32api.RegCloseKey(key)
	win32api.RegCloseKey(HKCU)
	
def SynchronizeRegistryWithDirectory(ADPropList):
	for Entry in config.ADProperties:
		AddRegistryKey(Entry,ADPropList[Entry])

def SetOutlookDefaults(WordCom):
	#If called and outlook is not open, it will open outlook with a gearbox icon. (rather not do this)
	#If called and outlook profile is not setup, all hell breaks loose. 
	EmailSignature = WordCom.EmailOptions.EmailSignature
	EmailSignature.NewMessageSignature = config.Settings['SignatureName']
	EmailSignature.ReplyMessageSignature = config.Settings['SignatureName']
	print3(EmailSignature.NewMessageSignature)
	print3(EmailSignature.ReplyMessageSignature)


def GetOutlookProcess():
	try:
		GetOutlookProcess.Counter
		GetOutlookProcess.Countdown
	except AttributeError:
		GetOutlookProcess.Counter = 5
		GetOutlookProcess.Countdown = False
	try:
		OutlookCom = win32com.client.GetActiveObject('Outlook.Application')
	except Exception as er:
		print4("Could not find outlook process")
		return False
	try:
		print2("Outlook Profile Found: %s" % OutlookCom.Application.DefaultProfileName)
	except Exception as er:
		GetOutlookProcess.Countdown = True
		print2("Outlook process found, but no profile found")
		return False
	if(not GetOutlookProcess.Countdown or GetOutlookProcess.Counter <= 0):
		print2("Outlook process found")
		return OutlookCom
	else:
		print2("Outlook and Profile found. Waiting %d more tries for stability" % GetOutlookProcess.Counter)
		GetOutlookProcess.Counter -=1
		return False

def WaitForOutlook():
	while True:
		OutlookCom = GetOutlookProcess()
		if not OutlookCom:
			time.sleep(5)
		else: 
			break
	return OutlookCom

def IsMAPIValid(mail):
	user = win32api.GetUserName()
	if('siggy' in mail.lower()): #ToDO Refactor this... It's coded specifically for the first client 
		return True 
	return False
	
def ConditionalReplace(DocuCom,SearchFor,Keep=True):
	Find = DocuCom.ActiveWindow.Selection.Find
	Find.MatchCase = False
	Find.MatchWholeWord = False
	Find.MatchWildcards = True
	Find.MatchSoundsLike = False
	Find.MatchAllWordForms = False
	Find.Forward = True
	Find.Wrap = 1 
	Find.Format = False
	Find.MatchKashida = False
	Find.MatchDiacritics = False
	Find.MatchAlefHamza = False
	Find.MatchControl = False
	SearchPattern = "\\<c_%s\\>(*)\\</c_%s\\>" % (SearchFor,SearchFor)
	if(Keep): ret = Find.Execute(SearchPattern,ReplaceWith='\\1',Replace=2)
	else: ret = Find.Execute(SearchPattern,ReplaceWith='',Replace=2)
	return ret
	
def WordReplace(DocuCom,SearchFor,ReplaceText):
	'''
	Params:
		MSWORD - MSWord COM Object
		SearchFor - What to search for (will be prefixed by Settings['ReplacementPrefix'])
		ReplaceText - What to Replace SearchFor with
	Returns:
		True on Success, False on Fail
	'''
	#https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.selection_members.aspx
	#https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.find_members.aspx
	#https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.replacement_members.aspx
	#Doctor = False
	ret = True
	#DocuCom.Visible = 0
	#DocuCom.DisplayAlerts = 0
	
	Find = DocuCom.ActiveWindow.Selection.Find
	Find.MatchCase = False
	Find.MatchWholeWord = False
	Find.MatchWildcards = False
	Find.MatchSoundsLike = False
	Find.MatchAllWordForms = False
	Find.Forward = True
	Find.Wrap = 1 
	Find.Format = False
	Find.MatchKashida = False
	Find.MatchDiacritics = False
	Find.MatchAlefHamza = False
	Find.MatchControl = False
	FindText = "<%s>" % SearchFor
	print("Finding %s Replacing %s" % (FindText,ReplaceText))
	if(DocuCom.ActiveWindow.Document.FullName == DocuCom.FullName):
		print3("ActiveDocument Fullname is a Match")
	else:
		print1("ActiveDocument %s is not a match..Exiting." % DocuCom.ActiveWindow.Document.FullName)
		sys.exit(0)
	if(not ReplaceText == None):
		try:
			ret = Find.Execute(FindText,ReplaceWith=ReplaceText,Replace=2)
			ConditionalReplace(DocuCom,SearchFor)
		except Exception as er:
			print1("Failed to modify Selection")
			ret = False
	else:
		try:
			ret = Find.Execute(FindText,ReplaceWith='',Replace=2)
			ConditionalReplace(DocuCom,SearchFor,False)
		except Exception as er:
			print1("Failed to modify Selection")
			ret = False
	return ret

def ClearDockLocks(SignatureFile=config.Settings['LocalSignature']):
	Active = True
	try:
		win32com.client.GetActiveObject("Word.Application")
	except Exception as er:
		print("[*] Word not active")
		Active = False 
	WordCom = win32com.client.gencache.EnsureDispatch("Word.Application")
	DocuCom = None

	print1("Active Object found")
	for Item in WordCom.Documents:
		if(Item.FullName == SignatureFile):
			print1("Local Signature already open in Active Object")
			DocuCom = Item
			print("Trace Set for Item in WordCom.Documents")
			return WordCom,DocuCom
	print1("Local Signature not already open in Active Object... Opening")
	DocuCom = WordCom.Documents.Open(SignatureFile,Visible=False)
	return WordCom,DocuCom,Active
		
def GenerateSignature(ADPropList):
	WordCom, DocuCom,Active = ClearDockLocks()
	Doctor=False
	if(not Active): WordCom.Application.Visible = 0
	#Dealing with Doctors
	if(ADPropList["displayname"].lower().startswith("dr.")):
		print("[*][*] Doctor Mode")
		Doctor=True
	#Done Dealing with Doctors
	wdFormatRTF = 6
	wdFormatHTML = 8
	wdFormatText = 2
	
	for Entry in config.ADProperties:
		#Dealing with Doctors
		if(Doctor and Entry == "givenname"):
			Value = "Dr. %s" % ADPropList["givenname"]
		else:
			Value = ADPropList[Entry]
		#Done Dealing with Doctors
		# Value = ADPropList[Entry] # Uncomment if you don't deal with doctors in a unique way like we do. 
		if(WordReplace(DocuCom,Entry,Value)):
			print1("Successful WordReplace on %s" % Value)
		else:
			print1("WordReplace Failed on %s" % Value)
	DocuCom.SaveAs("%s.txt" % config.Settings['OutlookSignature'], wdFormatText)
	DocuCom.SaveAs("%s.rtf" % config.Settings['OutlookSignature'], wdFormatRTF)
	DocuCom.SaveAs("%s.htm" % config.Settings['OutlookSignature'], wdFormatHTML)
	SetOutlookDefaults(WordCom)
	DocuCom.Close()
	if(not Active): WordCom.Quit()
	
def CheckFileSanity():
	ret = True # Sane
	if(not os.path.exists(config.Settings['LocalDirectory'])):
		try:
			os.mkdir(config.Settings['LocalDirectory'])
		except Exception as er:
			print1("Error, could not create %s" % config.Settings['LocalDirectory'])
			sys.exit(0)
	if(os.path.exists(config.Settings['MasterTemplate'])):  #Check Master Template
		print2('MasterTemplate found')

	else:
		print2("No master Template Found")
		sys.exit(0)
	return ret	
	
def QuickHash(filename):
	if(os.path.exists(filename)):
		hasher = hashlib.md5()
		with open(filename, 'rb') as hashbrown:
			buf = hashbrown.read(65536)
			while len(buf) > 0:
				hasher.update(buf)
				buf = hashbrown.read(65536)
		return hasher.hexdigest()
	
def IsSignatureUpdateNeeded():
	ret = False
	print2("Checking Signature Versions")
	#IS the signature there in the first place
	if(not os.path.exists(config.Settings['LocalSignature'])): 
		print2("No Signature found, copying")
		try:
			shutil.copyfile(config.Settings['MasterTemplate'],config.Settings['LocalSignature'])
		except Exception as er:
			print1("Could not copy signature")
			sys.exit(0)
		ret = True
	MasterTemplate = QuickHash(config.Settings['MasterTemplate'])
	LocalSignature = QuickHash(config.Settings['LocalSignature'])
	if (not MasterTemplate == LocalSignature):
		print2("Local and Master Signatures do not match")
		try:
			os.remove(config.Settings['LocalSignature'])
			shutil.copyfile(config.Settings['MasterTemplate'],config.Settings['LocalSignature'])
		except Exception as er:
			print1("Could not copy signature")
			sys.exit(0)
		ret = True
	return ret
	
def UpdateSignature(ADPropList):
	print1("UpdateSignature [*]: Waiting on Outlook")
	OutlookCom = WaitForOutlook()
	GenerateSignature(ADPropList)
	notify.BalloonTip(config.Settings['NotifyMessage'][0], config.Settings['NotifyMessage'][1])

def OverrideSignature():
	print1("Using Override Signature. Refreshing")
	OutlookCom = WaitForOutlook()
	WordCom, DocuCom, Active = ClearDockLocks(config.Settings['OverrideSignature'])
	wdFormatRTF = 6
	wdFormatHTML = 8
	wdFormatText = 2
	DocuCom.SaveAs("%s.txt" % config.Settings['OutlookSignature'], wdFormatText)
	DocuCom.SaveAs("%s.rtf" % config.Settings['OutlookSignature'], wdFormatRTF)
	DocuCom.SaveAs("%s.htm" % config.Settings['OutlookSignature'], wdFormatHTML)
	SetOutlookDefaults(WordCom)
	print1("Done Override")
	HKCU = win32api.RegOpenCurrentUser()
	win32api.RegDeleteTree(HKCU,config.Settings['RegistryLocation'])
	DocuCom.Close()
	if(not Active): WordCom.Quit()
	
def main():
	CheckFileSanity()
	RegPropList = GetRegistryInformation()
	ADPropList = GetDirectoryInformation()
	UpdateNeeded = False
	if(os.path.exists(config.Settings['OverrideSignature'])):
		OverrideSignature()
		sys.exit(0)
	if(not ADPropList):
		print("Could not Connect to AD")
		sys.exit(0)
	if(not IsDirectorySynchronized(RegPropList,ADPropList)):
		UpdateNeeded = True
		print1("Directory Mismatch")
		print1("Set Update Needed")
		SynchronizeRegistryWithDirectory(ADPropList)
	else:
		print1("Directory already Synchronized")
	if(IsSignatureUpdateNeeded()):
		print1("We need a signature update")
		UpdateNeeded = True
	else:
		print1("Signature Version is up to date")
	if(not IsMAPIValid(ADPropList['mail'])):
		print1("Not a valid email address, exiting")
		sys.exit(0)
	if(UpdateNeeded):
		print1("Updating Signature")
		UpdateSignature(ADPropList)
	
if __name__ == "__main__":
	main()
	
