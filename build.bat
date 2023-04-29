


go build -v -o ppt-summary.exe

del /Q %GOPATH%\bin\ppt-summary.exe
move ppt-summary.exe %GOPATH%\bin\ppt-summary.exe