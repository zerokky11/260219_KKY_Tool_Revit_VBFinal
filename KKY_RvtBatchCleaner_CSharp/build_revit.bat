@echo off
setlocal

if "%~1"=="" (
  set REVIT_YEAR=2019
) else (
  set REVIT_YEAR=%~1
)

echo Building for Revit %REVIT_YEAR%...
msbuild KKY_RvtBatchCleaner_CSharp.csproj /t:Build /p:Configuration=Release /p:RevitYear=%REVIT_YEAR%

if errorlevel 1 (
  echo.
  echo Build failed.
  exit /b 1
)

echo.
echo Build succeeded.
echo DLL / ADDIN deployed to %%APPDATA%%\Autodesk\Revit\Addins\%REVIT_YEAR%
exit /b 0
