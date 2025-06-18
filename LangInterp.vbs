Set Shell = CreateObject("WScript.Shell") 
Set Application = CreateObject("Shell.Application")
Set FileObject = CreateObject("Scripting.FileSystemObject")

' -| DateGrab Function |-------------------------------------------------------------
Function DateGrab()
    Dim today, year, month, day
    Dim Show

    today = Date
    year = DatePart("yyyy", today)
    month = DatePart("m", today)
    day = DatePart("d", today)

    Show = CStr(year) & "/" & Right("0" & CStr(month), 2) & "/" & Right("0" & CStr(day), 2)

    DateGrab = Show ' Pass the String to main...Or nothing will happen.
End Function

Dim CurrentLogPath

' -| CreateLog Function |------------------------------------------------------------
Sub CreateLog()
    Dim folderPath, baseName, ext, logPath, logNum
    folderPath = FileObject.GetParentFolderName(WScript.ScriptFullName)
    baseName = "log"
    ext = ".txt"
    logNum = 0
    
    Do
        If logNum = 0 Then
            logPath = FileObject.BuildPath(folderPath, baseName & ext)
        Else
            logPath = FileObject.BuildPath(folderPath, baseName & CStr(logNum) & ext)
        End If
        
        If Not FileObject.FileExists(logPath) Then
            Exit Do
        End If
        logNum = logNum + 1
    Loop
    
    Dim logFile
    Set logFile = FileObject.CreateTextFile(logPath, True)
    logFile.WriteLine "SATR Log File"
    logFile.Close
    
    CurrentLogPath = logPath
End Sub

' -| WriteToLog Function |-----------------------------------------------------------
Sub WriteToLog(text)
    Dim logFile
    If IsEmpty(CurrentLogPath) Or CurrentLogPath = "" Then
        MsgBox "CreateLog is required to be above WriteToLog", vbExclamation, "Fatal Error"
        Exit Sub
    End If
    
    Set logFile = FileObject.OpenTextFile(CurrentLogPath, 8, True)
    logFile.WriteLine Now & " - " & text
    logFile.Close
End Sub

' -| Secure Function |---------------------------------------------------------------
Function GenerateToken(length)
    Dim chars, i, token
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    Randomize
    token = ""
    For i = 1 To length
        token = token & Mid(chars, Int(Rnd() * Len(chars)) + 1, 1)
    Next
    GenerateToken = token
End Function

Class SecureFlag
    Private allowCreate
    Private internalToken
    Private bypassAntiTamper
    Private trustedCaller

    Private Sub Class_Initialize()
        internalToken = GenerateToken(64)
        allowCreate = False
        bypassAntiTamper = False
        trustedCaller = "langinterp.vbs"
    End Sub

    Private Function IsCalledFromTrustedScript()
        IsCalledFromTrustedScript = (LCase(WScript.ScriptName) = trustedCaller)
    End Function

    Public Property Get Token()
        Token = internalToken
    End Property

    Public Property Get CanCreate()
        CanCreate = allowCreate
    End Property

    Public Sub InternalSetAllowCreate(val, token)
        If token = internalToken Then
            If Not IsCalledFromTrustedScript() And Not bypassAntiTamper Then
                MsgBox "Warning: Direct tampering with a Security Value was Detected" & vbNewLine & _
                       "Execution will Halt.", vbCritical, "Security Warning"
                WScript.Quit
                Exit Sub
            End If
            allowCreate = val
        Else
            MsgBox "Security Warning: Invalid token used to set flag." & vbNewLine & _
                   "Execution will halt.", vbCritical, "SecureFlag Tampering Detected"
            WScript.Quit
        End If
    End Sub

    Public Sub SetBypassAllowCreation(flag, token)
        If token = internalToken Then
            bypassAntiTamper = flag
        End If
    End Sub
End Class

Dim gSecureFlag
Set gSecureFlag = New SecureFlag

Dim secureToken
secureToken = gSecureFlag.Token

Dim userConsentGiven
userConsentGiven = False

' -| VBSBatonPass Function |--------------------------------------------------------
Sub VBSBatonPass(filename, scriptContent)
    If Not gSecureFlag.CanCreate Then
        If Not userConsentGiven Then
            Dim ans
            ans = MsgBox("The current VBScript is attempting to create a new VBS file to run." & vbNewLine & _
                "This may be a potential security risk. Allow?", vbYesNo + vbExclamation, "Security Warning")

            If ans = vbYes Then
                gSecureFlag.SetBypassAllowCreation True, secureToken
                gSecureFlag.InternalSetAllowCreate True, secureToken
                gSecureFlag.SetBypassAllowCreation False, secureToken

                userConsentGiven = True
            Else
                MsgBox "Permission not granted, script will stop.", vbInformation, "Cancelled"
                Exit Sub
            End If
        Else
            MsgBox "Permission not granted to create new VBS files.", vbCritical, "Denied"
            Exit Sub
        End If
    End If

    Dim f
    Set f = FileObject.CreateTextFile(filename, True)
    f.Write scriptContent
    f.Close

    Shell.Run "wscript.exe """ & filename & """", 1, False
End Sub

' -| DeleteBatonPass Function |------------------------------------------------------
Sub DeleteBatonPass(filename)
    On Error Resume Next
    If FileObject.FileExists(filename) Then
        FileObject.DeleteFile filename, True
    End If
    If Err.Number <> 0 Then
        MsgBox "Failed to delete file: " & filename & vbNewLine & "Error: " & Err.Description, vbExclamation, "Delete Error"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' -| ReturnDirectory Function |-------------------------------------------------------
Function ReturnDirectory()
    On Error Resume Next
    Dim fullPath, folderPath
    fullPath = WScript.ScriptFullName
    folderPath = FileObject.GetParentFolderName(fullPath)
    ReturnDirectory = folderPath
    If Err.Number <> 0 Then
        MsgBox "Error retrieving current directory: " & Err.Description, vbExclamation, "GetCurrentDirectory Error"
        Err.Clear
    End If
    On Error GoTo 0
End Function