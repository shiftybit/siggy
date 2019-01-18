#To include AD Properties, Place them Here
import os
global ADProperties,Settings
ADProperties = list()
ADProperties.append("displayname")
ADProperties.append("mail")
ADProperties.append("title")
ADProperties.append("description")
ADProperties.append("telephonenumber")
ADProperties.append("facsimiletelephonenumber")
ADProperties.append("mobile")
ADProperties.append("streetaddress")
ADProperties.append("l") # City
ADProperties.append("givenname") # First Name
ADProperties.append("sn") # Last Name
ADProperties.append("initials") 
ADProperties.append("postalcode")  
ADProperties.append("physicaldeliveryofficename")  # Department
ADProperties.append("st")  # State
# ADProperties.append("professionalcredentials") # Removed. This was from an environment where I had extended the AD Schema to accommodate the need
#ADProperties.append("signaturequote")    # Removed. This was from an environment where I had extended the AD Schema to accommodate the need


#Coded for HKCU. 
Settings = dict()
Settings['RegistryLocation'] = "Siggy"
Settings['SignatureName'] = "Siggy_Standard_Signature" 
Settings['OutlookSignaturePath'] = "%s\\Microsoft\\Signatures" % os.getenv('APPDATA')
Settings['OutlookSignature'] = "%s\\%s" % (Settings['OutlookSignaturePath'],Settings['SignatureName'])
Settings['RunningDirectory'] = os.getcwd()
Settings['LocalDirectory'] = os.path.join(os.getenv('APPDATA'),"siggy")
Settings['LocalSignature'] = os.path.join(Settings['LocalDirectory'],"%s.docx" % Settings['SignatureName'])
Settings['OverrideSignature'] = os.path.join(Settings['LocalDirectory'],"override.docx")
Settings['MasterTemplate'] = "%s\\Signature\\%s.docx" % (Settings['RunningDirectory'],Settings['SignatureName'])
Settings['NotifyMessage'] = "Microsoft Outlook", "Your Email Signature has been Updated"