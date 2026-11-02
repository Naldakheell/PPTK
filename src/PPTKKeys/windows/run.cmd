@echo off
setlocal
cd /d "%~dp0"
dotnet run --project PPTKKeys.csproj -c Release -- %*
exit /b %errorlevel%
