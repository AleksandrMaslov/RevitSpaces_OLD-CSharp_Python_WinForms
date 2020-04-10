@echo deleting addin file and plugin folder

set addinDir="C:\ProgramData\Autodesk\Revit\Addins\2019"
set delFile="Synsys_CreateSpacesFromLinkedRooms.addin"
set delDir=%addinDir%\SynSys\CreateSpacesFromLinkedRooms

if exist %addinDir%\%delFile% (
del %addinDir%\%delFile%
)

if exist %delDir% (
@RD /S /Q %delDir%
)

pause