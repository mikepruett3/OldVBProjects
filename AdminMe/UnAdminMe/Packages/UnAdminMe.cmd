@ECHO OFF
setlocal
set _USER_=%USERNAME%
set _ADMIN_=Administrator
set _ADMING_=Administrators
set _SYS32_=%SYSTEMROOT%\System32

set _PROG_=%_SYS32_%\schtasks /DELETE /TN UnAdminMe /F
ECHO %_PROG_%
%_PROG_%
REM runas /u:%_ADMIN_% %_PROG_%"

set _PROG_=
set _PROG_=NET LOCALGROUP %_ADMING_% %_USER_% /DELETE
%_PROG_%
REM runas /u:%_ADMIN_% %_PROG_%

