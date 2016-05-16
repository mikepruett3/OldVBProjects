Attribute VB_Name = "Module1"

Public Sub Main()
    Dim strKeyPath As String
    Dim strComputer As String
    Dim strEntryName As String
    Dim dwValue As String
    Dim objReg As Object
    On Error Resume Next
    Const HKEY_LOCAL_MACHINE = &H80000002
    strComputer = "."
    strKeyPath = "Software\Policies\Microsoft\Windows\WindowsUpdate"
    strEntryName = "DoNotAllowXPSP2"
    dwValue = 1
    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
    AddBlock
    Err.Clear
End Sub

Sub AddBlock() 'Check whether WindowsUpdate subkey exists.
    Dim strParentPath As String
    Dim strTargetSubKey As String
    Dim intCount As Integer
    Dim intReturn1 As String
    strParentPath = "SOFTWARE\Policies\Microsoft\Windows"
    strTargetSubKey = "WindowsUpdate"
    intCount = 0
    intReturn1 = objReg.EnumKey(HKEY_LOCAL_MACHINE, strParentPath, arrSubKeys)
    'MsgBox intReturn1
    If intReturn1 = 0 Then
        For Each strSubKey In arrSubKeys
            MsgBox strSubKey
            If strSubKey = strTargetSubKey Then
                intCount = 1
            End If
        Next
        If intCount = 1 Then
            SetValue
        Else
            MsgBox ("Unable to find registry subkey " & _
            strTargetSubKey & ". Creating ...")
            intReturn2 = objReg.CreateKey(HKEY_LOCAL_MACHINE, _
            strKeyPath)
            If intReturn2 = 0 Then
                SetValue
            Else
                MsgBox ("ERROR: Unable to create registry " & _
                "subkey " & strTargetSubKey & ".")
            End If
        End If
    Else
        MsgBox ("ERROR: Unable to find registry path " & _
        strParentPath & ".")
    End If
End Sub

Sub SetValue()
    intReturn = objReg.SetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath _
                , strEntryName, dwValue)
    If intReturn = 0 Then
        MsgBox ("Added registry entry to block Windows XP " & _
        "SP2 deployment via Windows Update or Automatic Update.")
    Else
        MsgBox ("ERROR: Unable to add registry entry to " & _
        "block Windows XP SP2 deployment via Windows Update " & _
        "or Automatic Update.")
    End If
End Sub
