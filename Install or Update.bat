@CLS
@ECHO OFF
COLOR 0A
TITLE Update Inspector Mike Excel Addin

ECHO.  
ECHO Inspector Mike Excel Addin
ECHO.  

ECHO.
ECHO  This will attempt to copy InspectorMike_Addin.xlam to %appdata%\Microsoft\Addins
ECHO  and this will fail if Microsoft Excel is open.
ECHO.
ECHO  Please close Microsoft Excel first.
ECHO.
ECHO.  

DIR %appdata%\Microsoft\Addins

SET CONTINUE=N
SET /P CONTINUE=Are you ready to copy the file (Y/N)?[%CONTINUE%]: 

:: WOW - meet boolean OR in DOSland

IF "%CONTINUE%"=="Y" GOTO COPY
IF "%CONTINUE%"=="y" GOTO COPY

:NOCOPY
	ECHO.
	ECHO No file copied.
	ECHO.

	GOTO PAUSE

:COPY
	ECHO.
	COPY InspectorMike_Addin.xlam %appdata%\Microsoft\Addins
	ECHO.

	DIR %appdata%\Microsoft\Addins
	

	GOTO PAUSE

:PAUSE
	Pause
