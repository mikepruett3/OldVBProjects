Attribute VB_Name = "Module1"
Public Sub Main()
    Dim objShell As Object
    Dim KeyOne As String
    Dim KeyTwo As String
    Dim ISTAKey As String
    Dim strComputer As String
    Dim objWMIService As Object
    Dim objNetwork As Object
    Dim colComputers As Object
    
    Set objNetwork = CreateObject("WScript.Network")
    strComputer = objNetwork.ComputerName
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colComputers = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")

    For Each objComputer In colComputers
        Select Case objComputer.DomainRole
            Case 0
                strComputerRole = "Standalone Workstation"
                ApplyKeys
            Case 1
                strComputerRole = "Member Workstation"
                ApplyKeys
            Case 2
                strComputerRole = "Standalone Server"
                ApplyKeys
            Case 3
                strComputerRole = "Member Server"
                ApplyKeys
            Case 4
                strComputerRole = "Backup Domain Controller"
                WScript.Echo "Cannot Apply to Backup Domain Controllers."
            Case 5
                strComputerRole = "Primary Domain Controller"
                WScript.Echo "Cannot Apply to Primary Domain Controllers."
        End Select
    Next
End Sub

Sub ApplyKeys()
    Dim objShell As Object
    Dim strKeyTop As String
    Dim strKeyOne As String
    Dim strKeyTwo As String
    Dim strKeyTemp As String
    strKeyTop = "HKLM\SYSTEM\CurrentControlSet\Services\Browser\Parameters\"
    strKeyOne = "IsDomainMaster"
    strKeyTwo = "MaintainServerList"
    Set objShell = CreateObject("WScript.Shell")
    On Error Resume Next
    strKeyTemp = objShell.RegRead(strKeyTop & strKeyOne)
    If (strKeyTemp <> "FALSE") Then
        objShell.RegWrite strKeyTop & KeyOne, "FALSE", "REG_SZ"
    End If
    strKeyTemp = ""
    strKeyTemp = objShell.RegRead(strKeyTop & strKeyTwo)
    If (strKeyTemp <> "FALSE") Then
        objShell.RegWrite strKeyTop & strKeyTwo, "FALSE", "REG_SZ"
    End If
End Sub
