@echo off
setlocal
for /f "tokens=1,2*" %%i in ('reg query HKLM\software\Microsoft\VisualStudio\8.0 /v InstallDir') do set VSDir=%%k
"%VSDIR%\devenv.exe" %*
