Attribute VB_Name = "Module1"
Option Explicit

Private Const CREATE_DEFAULT_ERROR_MODE = &H4000000

Private Const LOGON_WITH_PROFILE = &H1
Private Const LOGON_NETCREDENTIALS_ONLY = &H2

Private Const LOGON32_LOGON_INTERACTIVE = 2
Private Const LOGON32_PROVIDER_DEFAULT = 0

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long ' !!! must be Long for Unicode string
    lpDesktop As Long  ' !!! must be Long for Unicode string
    lpTitle As Long    ' !!! must be Long for Unicode string
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

'  LogonUser() requires that the caller has the following permission
'  Permission                        Display Name
'  --------------------------------------------------------------------
'  SE_TCB_NAME                      Act as part of the operating system

'  CreateProcessAsUser() requires that the caller has the following permissions
'  Permission                        Display Name
'  ---------------------------------------------------------------
'  SE_ASSIGNPRIMARYTOKEN_NAME       Replace a process level token
'  SE_INCREASE_QUOTA_NAME           Increase quotas

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
      hHandle As Long, ByVal dwMilliseconds As Long) As Long
      
Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
      lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" _
      (ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
      (ByVal hProcess As Long, lpExitCode As Long) As Long
      

Private Declare Function LogonUser Lib "advapi32.dll" Alias _
        "LogonUserA" _
        (ByVal lpszUsername As String, _
        ByVal lpszDomain As String, _
        ByVal lpszPassword As String, _
        ByVal dwLogonType As Long, _
        ByVal dwLogonProvider As Long, _
        phToken As Long) As Long

Private Declare Function CreateProcessAsUser Lib "advapi32.dll" _
        Alias "CreateProcessAsUserA" _
        (ByVal hToken As Long, _
        ByVal lpApplicationName As Long, _
        ByVal lpCommandLine As String, _
        ByVal lpProcessAttributes As Long, _
        ByVal lpThreadAttributes As Long, _
        ByVal bInheritHandles As Long, _
        ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, _
        ByVal lpCurrentDirectory As String, _
        lpStartupInfo As STARTUPINFO, _
        lpProcessInformation As PROCESS_INFORMATION) As Long

' CreateProcessWithLogonW API is available only on Windows 2000 and later.
Private Declare Function CreateProcessWithLogonW Lib "advapi32.dll" _
        (ByVal lpUsername As String, _
        ByVal lpDomain As String, _
        ByVal lpPassword As String, _
        ByVal dwLogonFlags As Long, _
        ByVal lpApplicationName As Long, _
        ByVal lpCommandLine As String, _
        ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, _
        ByVal lpCurrentDirectory As String, _
        ByRef lpStartupInfo As STARTUPINFO, _
        ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
                                  
Private Declare Function SetErrorMode Lib "kernel32.dll" _
        (ByVal uMode As Long) As Long
        
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
                             
' Version Checking APIs
Private Declare Function GetVersionExA Lib "kernel32.dll" _
    (lpVersionInformation As OSVERSIONINFO) As Integer

Private Const VER_PLATFORM_WIN32_NT = &H2

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Function ExecCmd(cmdline$)
      Dim proc As PROCESS_INFORMATION
      Dim start As STARTUPINFO, strReturn As Variant

      ' Initialize the STARTUPINFO structure:
      start.cb = Len(start)

      ' Start the shelled application:
      strReturn = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, _
         NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)

      ' Wait for the shelled application to finish:
         strReturn = WaitForSingleObject(proc.hProcess, INFINITE)
         Call GetExitCodeProcess(proc.hProcess, strReturn)
         Call CloseHandle(proc.hThread)
         Call CloseHandle(proc.hProcess)
         ExecCmd = strReturn
   End Function

'********************************************************************

'                   RunAsUser for Windows 2000 and Later
'********************************************************************
Public Function W2KRunAsUser(ByVal UserName As String, _
        ByVal Password As String, _
        ByVal DomainName As String, _
        ByVal CommandLine As String, _
        ByVal CurrentDirectory As String) As Long

    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    
    Dim wUser As String
    Dim wDomain As String
    Dim wPassword As String
    Dim wCommandLine As String
    Dim wCurrentDir As String
    
    Dim Result As Long
    
    si.cb = Len(si)
        
    wUser = StrConv(UserName + Chr$(0), vbUnicode)
    wDomain = StrConv(DomainName + Chr$(0), vbUnicode)
    wPassword = StrConv(Password + Chr$(0), vbUnicode)
    wCommandLine = StrConv(CommandLine + Chr$(0), vbUnicode)
    wCurrentDir = StrConv(CurrentDirectory + Chr$(0), vbUnicode)
    
    Result = CreateProcessWithLogonW(wUser, wDomain, wPassword, _
          LOGON_WITH_PROFILE, 0&, wCommandLine, _
          CREATE_DEFAULT_ERROR_MODE, 0&, wCurrentDir, si, pi)
    ' CreateProcessWithLogonW() does not
    If Result <> 0 Then
        CloseHandle pi.hThread
        CloseHandle pi.hProcess
        W2KRunAsUser = 0
    Else
        W2KRunAsUser = Err.LastDllError
        MsgBox "CreateProcessWithLogonW() failed with error " & Err.LastDllError, vbExclamation
    End If

End Function

'********************************************************************
'                   RunAsUser for Windows NT 4.0
'********************************************************************
Public Function NT4RunAsUser(ByVal UserName As String, _
                ByVal Password As String, _
                ByVal DomainName As String, _
                ByVal CommandLine As String, _
                ByVal CurrentDirectory As String) As Long
Dim Result As Long
Dim hToken As Long
Dim si As STARTUPINFO
Dim pi As PROCESS_INFORMATION

    Result = LogonUser(UserName, DomainName, Password, LOGON32_LOGON_INTERACTIVE, _
                       LOGON32_PROVIDER_DEFAULT, hToken)
    If Result = 0 Then
        NT4RunAsUser = Err.LastDllError
        ' LogonUser will fail with 1314 error code, if the user account associated
        ' with the calling security context does not have
        ' "Act as part of the operating system" permission
        MsgBox "LogonUser() failed with error " & Err.LastDllError, vbExclamation
        Exit Function
    End If
    
    si.cb = Len(si)
    Result = CreateProcessAsUser(hToken, 0&, CommandLine, 0&, 0&, False, _
                CREATE_DEFAULT_ERROR_MODE, _
                0&, CurrentDirectory, si, pi)
    If Result = 0 Then
        NT4RunAsUser = Err.LastDllError
        ' CreateProcessAsUser will fail with 1314 error code, if the user
        ' account associated with the calling security context does not have
        ' the following two permissions
        ' "Replace a process level token"
        ' "Increase Quotoas"
        MsgBox "CreateProcessAsUser() failed with error " & Err.LastDllError, vbExclamation
        CloseHandle hToken
        Exit Function
    End If
    
    CloseHandle hToken
    CloseHandle pi.hThread
    CloseHandle pi.hProcess
    NT4RunAsUser = 0

End Function

Public Function RunAsUser(ByVal UserName As String, _
                ByVal Password As String, _
                ByVal DomainName As String, _
                ByVal CommandLine As String, _
                ByVal CurrentDirectory As String) As Long

    Dim w2kOrAbove As Boolean
    Dim osinfo As OSVERSIONINFO
    Dim Result As Long
    Dim uErrorMode As Long
    
    ' Determine if system is Windows 2000 or later
    osinfo.dwOSVersionInfoSize = Len(osinfo)
    osinfo.szCSDVersion = Space$(128)
    GetVersionExA osinfo
    w2kOrAbove = _
        (osinfo.dwPlatformId = VER_PLATFORM_WIN32_NT And _
         osinfo.dwMajorVersion >= 5)
    If (w2kOrAbove) Then
        Result = W2KRunAsUser(UserName, Password, DomainName, _
                    CommandLine, CurrentDirectory)
    Else
        Result = NT4RunAsUser(UserName, Password, DomainName, _
                    CommandLine, CurrentDirectory)
    End If
    RunAsUser = Result
End Function
Public Sub Pause(Seconds As Single)
    Dim T As String, T2 As Single, Num As Single
    Num = Seconds * 1000
    T = GetTickCount()
    T2 = GetTickCount()
    Do Until T2 - T >= Num
        T2 = GetTickCount()
        DoEvents
    Loop
End Sub
Public Function TimeToSleep(Min As String)
    Dim strTime As Single
    strTime = Min * 60
    TimeToSleep = strTime
End Function
Public Sub RunProg(strRun As String) ', ByVal sender As System.Object, ByVal e As System.EventArgs)
    'System.Diagnostics.Process.Start (strRun)
    Shell strRun, vbNormalFocus
End Sub

Public Sub Main()
    Dim strUser As String, strPassword As String, strWorkingDir As String
    Dim strCommand As String, strProgram As String, strDrive As String
    Dim strWin As String, strHost As String, strPrgFil As String, strUserName As String
    Dim strAdminGrp As String, strMinutes As String, strSeconds As String
    Dim ExitCode As String, oWShell As Object, strDomain As String
    
    Set oWShell = CreateObject("WScript.Shell")
    strDrive = oWShell.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
    strWin = oWShell.ExpandEnvironmentStrings("%SYSTEMROOT%")
    strHost = oWShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    strPrgFil = oWShell.ExpandEnvironmentStrings("%PROGRAMFILES%")
    strUserName = oWShell.ExpandEnvironmentStrings("%USERNAME%")
    strDomain = oWShell.ExpandEnvironmentStrings("%USERDOMAIN%")
    strUser = "Administrator"
    strPassword = ""
    strAdminGrp = "Administrators"
    strWorkingDir = strWin & "\system32\"
    strMinutes = InputBox("How long will you need Elevated Access?" & vbCrLf & vbCrLf & _
                       "(In Whole Minutes Please...", "How Long?", "15")
    strSeconds = TimeToSleep(strMinutes)
    strProgram = "NET LOCALGROUP " & strAdminGrp & " " & Chr(34) & strDomain & "\" & strUserName & Chr(34) & " /ADD"
    RunAsUser strUser, strPassword, strHost, strProgram, strWorkingDir
    strProgram = "NET STOP " & Chr(34) & "Cisco Security Agent" & Chr(34)
    RunAsUser strUser, strPassword, strHost, strProgram, strWorkingDir
    Pause (35)
    strProgram = "pskill explorer.exe"
    ExecCmd (strProgram)
    Pause (35)
    strProgram = "psexec.exe -d -u " & strUserName & " explorer.exe"
    ExecCmd (strProgram)
    Pause (35)
    strProgram = "REGEDIT /S " & strWorkingDir & "l_amdrem.reg"
    RunAsUser strUser, strPassword, strHost, strProgram, strWorkingDir
    strProgram = "NET LOCALGROUP " & strAdminGrp & " " & Chr(34) & strDomain & "\" & strUserName & Chr(34) & " /DELETE"
    RunAsUser strUser, strPassword, strHost, strProgram, strWorkingDir
    strProgram = "shutdown -r -t " & strSeconds & " -f -c " & Chr(34) & "This Program will cause your system to reboot after the timer reaches Zero. Please save all open documents before then." & Chr(34)
    RunAsUser strUser, strPassword, strHost, strProgram, strWorkingDir
End Sub

