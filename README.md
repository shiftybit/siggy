# siggy
Siggy is a per user Group Policy Logon Executable. It updates email signatures based on a Word Document and Data about the current user in active directory. 
Siggy does not require python to be installed on the client machines. It is packaged into a single executable via pyInstaller. 

Tested on python 3.15. Other versions have issues with the win32api. 

Setup:
Edit the Signature-Standard.docx in the Signatures folder. 
Deploy signature standard to a network location. 
edit config.py and point the variables to the correct UNC path. Ensure users have access to this.  (or you can replicate via Group Policy file preferences) 
use pyinstaller to build siggy.py
set siggy.exe as a logon script.

Requirements:
Active Directory
Microsoft Office (Tested on 2010 and up)
