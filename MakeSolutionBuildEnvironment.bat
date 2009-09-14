regsvr32 /u /s Bin\SolutionBuildEnvironment.dll

setlocal
for /f "tokens=1,2*" %%i in ('reg query HKLM\software\Microsoft\VisualStudio\9.0 /v InstallDir') do set VSDir=%%k
"%VSDIR%\devenv.com" SolutionBuildEnvironment.sln /rebuild Release
endlocal

setlocal
for /f "tokens=1,2,3,4,5*" %%i in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 5_is1" /v "Inno Setup: App Path"') do set INNO=%%n
"%INNO%\iscc" SolutionBuildEnvironment.iss
endlocal

mkdir Out
zip SolutionBuildEnvironment130.zip SolutionBuildEnvironment130.exe
move SolutionBuildEnvironment130.exe Out
move SolutionBuildEnvironment130.zip Out
copy *.html Out
del Addin.h
del Addin_i.c
del *.aps
del *.bak
del *.ncb
del *.WW
zip SolutionBuildEnvironment130_Source * Bin\SolutionBuildEnvironment.dll
move SolutionBuildEnvironment130_Source.zip Out

