Attribute VB_Name = "Module1"
Sub Main()
    Dim objShell As Object
    Dim strKey1 As String
    Dim strKeyTmp As String
    On Error Resume Next
    strKey1 = "HKLM\SOFTWARE\Microsoft\Exchange\Client\Options\DumpsterAlwaysOn"
    Set objShell = CreateObject("WScript.Shell")
    strKeyTemp = objShell.RegRead(strKey1)
    If (strKeyTemp = "") Or (strKeyTemp = 0) Then ' Key Missing, then Create!
        objShell.RegWrite strKey1, "00000001", "REG_DWORD"
    End If
End Sub
