Attribute VB_Name = "codemain"
Public Sub Main()
    Dim DomainName As String
    Dim UserAccount As String
    Dim strLocal As String
    Dim net As Object
    Dim OLE As Object
    Dim group As Object
        
    Set net = CreateObject("WScript.Network")
    strLocal = net.ComputerName
    'MsgBox strLocal
    DomainName = InputBox("Enter Domain Name:")
    UserAccount = InputBox("Enter User Name:")
    
    Set group = GetObject("WinNT://" & strLocal & "/Administrators")
    group.Filter = Array("group")
    
    For Each domuser In group
        MsgBox domuser.Names
    Next
    
    'On Error Resume Next
    'group.Delete "WinNt://" & DomainName & "/" & UserAccount & ""
    'If Not Err.Number = 0 Then
    '   Set OLE = CreateObject("ole.err")
    '   MsgBox OLE.oleError(Err.Number), vbCritical
    '   Err.Clear
    'Else
    '    MsgBox "Done."
    'End If
End Sub
