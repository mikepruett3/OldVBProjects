@ECHO OFF
setlocal
set _USER_=%USERNAME%
set _ADMIN_=Administrator
set _ADMING_=Administrators
set _PROG_="NET LOCALGROUP %_ADMING_% %_USER_% /ADD"
set _SYS32_=%SYSTEMROOT%\System32

runas /u:%_ADMIN_% %_PROG_%

set _PROG_=
set _PROG_="%_SYS32_%\schtasks /CREATE /SC minute /MO 5 /TR %_SYS32_%\UnAdminMe.cmd /TN UnAdminMe /RU "SYSTEM""
REM ECHO %_PROG_%
runas /u:%_ADMIN_% %_PROG_%"
