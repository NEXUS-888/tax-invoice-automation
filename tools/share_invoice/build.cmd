@echo off
rem Builds ShareInvoice.exe with the C# compiler that ships with .NET Framework 4 (part of Windows 10/11).
setlocal
set NETFX=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
set WINMD=%WINDIR%\System32\WinMetadata
set OUT=%~dp0ShareInvoice.exe

"%NETFX%\csc.exe" /nologo /target:winexe /platform:anycpu /optimize+ /out:"%OUT%" ^
  /r:"%WINMD%\Windows.ApplicationModel.winmd" ^
  /r:"%WINMD%\Windows.Storage.winmd" ^
  /r:"%WINMD%\Windows.Foundation.winmd" ^
  /r:"%NETFX%\System.Runtime.dll" ^
  /r:"%NETFX%\System.Runtime.InteropServices.WindowsRuntime.dll" ^
  /r:"%NETFX%\System.Threading.Tasks.dll" ^
  /r:System.Windows.Forms.dll /r:System.Drawing.dll ^
  "%~dp0ShareInvoice.cs"
exit /b %ERRORLEVEL%
