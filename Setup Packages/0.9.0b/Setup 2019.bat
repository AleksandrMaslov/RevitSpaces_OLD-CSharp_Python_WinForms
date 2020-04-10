@echo on
for %%i in (%CD%) do set "name=%%~nxi"
cd Lib

set addinDir="C:\ProgramData\Autodesk\Revit\Addins\2019"

xcopy "%CD%\*.addin" %addinDir% /r /y
xcopy "%CD%" %addinDir%\SynSys\CreateSpacesFromLinkedRooms /s /i /r /y /EXCLUDE:exclude.txt

pause