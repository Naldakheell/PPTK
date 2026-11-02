@echo off
setlocal
cd /d "%~dp0"
where dotnet >nul 2>&1
if errorlevel 1 (
  echo Install the .NET 10 SDK from https://dotnet.microsoft.com/download/dotnet/10.0
  exit /b 1
)
dotnet restore PPTKKeys.csproj
if errorlevel 1 exit /b 1
dotnet build PPTKKeys.csproj -c Release --nologo
exit /b %errorlevel%
