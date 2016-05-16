Attribute VB_Name = "Module1"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMillseconds As Long)
Public Sub Main()
    Dim strUser As String
    Dim strPassword As String
    Dim strCommand As String
    Dim strProgram As String
    Dim strDrive As String
    Dim strWin As String
    Dim ExitCode As String
    Dim oWShell As Object
    'Dim oSleep As Object
    
    Set oWShell = CreateObject("WScript.Shell")
    'Set oSleep = CreateObject("WScript.Sleep")
    strDrive = oWShell.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
    strWin = oWShell.ExpandEnvironmentStrings("%SYSTEMROOT%")
    strHost = oWShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    strUser = "Administrator"
    strPassword = ""
    strProgram = "RunDLL32.exe shell32.dll,Control_RunDLL " & strWin & "\system32\timedate.cpl"
    
    On Error Resume Next
    strCommand = "runas.exe /env /user:" & strHost & "\" & strUser & " " & Chr(34) & strProgram & Chr(34)
    MsgBox strCommand
    ExitCode = wshShell.Run(strCommand, 1, True)
    Sleep 500
    'oSleep.Sleep 500
    oWShell.SendKeys strPassword
End Sub
