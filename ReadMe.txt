-------------------------------------
|	   SATR Sub-Language :P		  |
|          A VBScript Subsidiarity                  |  
|	          By: fhAifN				  |
-------------------------------------

Custom Stuff:

DateGrab:
Grab's the Current Date for the current User.
Doesn't Support Arguments.

CreateLog:
Create's a Log within the current Directory that the VBScript is running in.
Doesn't Support Arguments.

WriteToLog:
Writes to that Log created by CreateLog.
Doesn't Support Arguments. CreateLog must come before WriteToLog, or an Error will be raised.

VBSBatonPass:
Creates a Child Process/Child VBScript by a User Specified Dim. Supports vbNewLine and it's Aliases.
Supports only the ,NewScript argument. A Dim is required before VBSBatonPass.
VBSBatonPass uses Private Classes, and Private Random Keys to prevent an Unauthorized BatonPass.
VBSBatonPass also requires User Consent before a VBS is created and Ran.

DeleteBatonPass:
Delete's the VBS made by BatonPass.
Name of VBS file is required for a deletion to occur.

ReturnDirectory:
Returns the Current Directory.
Doesn't Support Arguments.