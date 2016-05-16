'Start of Script
'VBRUNAS.VBS
'v1.2 March 2001
'Jeffery Hicks
'jhicks@quilogy.com http://www.quilogy.com
'USAGE: cscript|wscript VBRUNAS.VBS Username Password Command
'DESC: A RUNAS replacement to take password at a command prompt.
'NOTES: This is meant to be used for local access. If you want to run a command
'across the network as another user, you must add the /NETONLY switch to the RUNAS 
'command.

' *********************************************************************************
' * THIS PROGRAM IS OFFERED AS IS AND MAY BE FREELY MODIFIED OR ALTERED AS *
' * NECESSARY TO MEET YOUR NEEDS. THE AUTHOR MAKES NO GUARANTEES OR WARRANTIES, *
' * EXPRESS, IMPLIED OR OF ANY OTHER KIND TO THIS CODE OR ANY USER MODIFICATIONS. *
' * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED IN A SECURED LAB *
' * ENVIRONMENT. USE AT YOUR OWN RISK. *
' *********************************************************************************

On Error Resume Next
dim WshShell,oArgs,FSO

set oArgs=wscript.Arguments

if InStr(oArgs(0),"?")<>0 then
wscript.echo VBCRLF & "? HELP ?" & VBCRLF
Usage
end if

if oArgs.Count <3 then
wscript.echo VBCRLF & "! Usage Error !" & VBCRLF
Usage
end if

sUser=oArgs(0)
WScript.Echo sUser
sPass=oArgs(1)& VBCRLF
WScript.Echo sPass
sCmd=oArgs(2)
WScript.Echo sCmd

set WshShell = CreateObject("WScript.Shell")
set WshEnv = WshShell.Environment("Process")
WinPath = WshEnv("SystemRoot")&"\System32\runas.exe"
set FSO = CreateObject("Scripting.FileSystemObject")

if FSO.FileExists(WinPath) then
	wscript.echo WinPath & " " & "verified"
else
	wscript.echo "!! ERROR !!" & VBCRLF & "Can't find or verify " & winpath &"." & VBCRLF & "You must be running Windows 2000 for this script to work."
	set WshShell=Nothing
	set WshEnv=Nothing
	set oArgs=Nothing
	set FSO=Nothing
	wscript.quit
end if

Dim strProg
strProg = "runas /user:" & sUser & " " & Chr(34) & sCmd & Chr(34)
WScript.Echo strProg
rc=WshShell.Run(strProg, 1, TRUE)
'Wscript.Sleep 30 'need to give time for window to open.
WshShell.AppActivate(WinPath) 'make sure we grab the right window to send password to
WshShell.SendKeys sPass 'send the password to the waiting window.

set WshShell=Nothing
set oArgs=Nothing
set WshEnv=Nothing
set FSO=Nothing

'wscript.quit

'************************
'* Usage Subroutine *
'************************
Sub Usage()
On Error Resume Next
msg="Usage: cscript|wscript vbrunas.vbs Username Password Command" & VBCRLF & VBCRLF & "You should use the full path where necessary and put long file names or commands" & VBCRLF & "with parameters in quotes" & VBCRLF & VBCRLF &"For example:" & VBCRLF &" cscript vbrunas.vbs quilogy\jhicks luckydog e:\scripts\admin.vbs" & VBCRLF & VBCRLF &" cscript vbrunas.vbs quilogy\jhicks luckydog " & CHR(34) &"e:\program files\scripts\admin.vbs 1stParameter 2ndParameter" & CHR(34)& VBCRLF & VBCRLF & VBCLRF & "cscript vbrunas.vbs /?|-? will display this message."

wscript.echo msg

wscript.quit

end sub
'End of Script 