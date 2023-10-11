MSBuild.exe "Toms Updater.sln" /noconsolelogger /t:Rebuild /p:Configuration=Release
cd "Toms Updater\bin\Release"
"Package This Program!.bat"