' -| THESE ARE REQUIRED!!! REMOVING THESE LINES WILL RESULT IN A ERROR. |--------------------------------------------------
Set Shell = CreateObject("WScript.Shell") 
Set Application = CreateObject("Shell.Application")
Set FileObject = CreateObject("Scripting.FileSystemObject")

Dim ts, Read
Set ts = FileObject.OpenTextFile("LangInterp.vbs")
Read = ts.ReadAll()
ts.Close

ExecuteGlobal Read
' ==========================================================================================================================
' -| Insert your now defined custom functions (or base VBS, doesn't matter) here! |------------------------------------------------------
