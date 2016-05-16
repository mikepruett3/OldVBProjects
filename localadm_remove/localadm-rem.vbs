' --------------------------------------------------------
' - Local_Administrators_Removal
' - Script: localadm-rem.vbs
' - Usage: wscript localadm-rem.vbs
' - Author: Mike Pruett
' - 		amanoj <at> gmail <dot> com
' - Created: March 9th, 2006
' - Revision: .009
' - Desc: This script was created to remove all users from
' - the Local Administrators group from each Workstation
' - on the selected domain. Make sure to change the "DomainName"
' - variable to reflect the target domain name. Then update the 
' - "strSafeUsers" array with those user accounts that should be
' - left in the group. This will not remove the "Administrator" 
' - account, as it will fail.
' --------------------------------------------------------
Dim DomainName,strLocal,strUsers,user ' As String
Dim objNetwork,objGroup,objUser,objOle ' AS Object

DomainName = "ISTAADS"
Set objNetwork = CreateObject("WScript.Network")
'strLocal = objNetwork.ComputerName
Set objGroup = GetObject("WinNT://"& DomainName & "/" & strLocal &"/Administrators,group")

On Error Resume Next
For each strUsers in objGroup.Members
   WScript.Echo strUsers.Name
   user = strUsers.Name
   If user = "jjauod" then
   	  Set objUser = GetObject("WinNT://" & DomainName & "/" & strLocal & "/" & user & ",user")
   	  objGroup.Remove(objUser.ADsPath)
   	  CheckError
   End If
Next

Sub CheckError
	If Not err.number=0 Then
		Set objOle = CreateObject("ole.err")
		MsgBox objOle.oleError(err.Number), vbCritical
		err.clear
	else
		MsgBox "Done."
	End If
End Sub