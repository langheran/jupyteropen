#Persistent
#SingleInstance off
DetectHiddenWindows, On

args:=""
Loop, %0%  ; For each parameter:
{
    param := %A_Index%  ; Fetch the contents of the variable whose name is contained in A_Index.
	num = %A_Index%
	args := args . param
}
if(args="")
	args:=A_WorkingDir
if(!FileExist(args))
	ExitApp

SplitPath, args, name, dir,,name_no_ext

command="C:\Windows\System32\bash.exe" -ilc "source activate xeus; jupyter notebook"
Run, %comspec% /k "%command%", %dir%
WinWaitActive %Comspec%
processTitle=%dir% - JupyterOpen
WinSetTitle %processTitle%
WinHide %processTitle%
Sleep, 2000
Menu, Tray, Tip, %dir%
Menu, Tray, NoStandard
Menu, Tray, Add, &Exit, saveAndExit
Menu, Tray, add, &Open Browser, OpenRoot
Menu, Tray, add, &Open Folder, OpenFolder
root:="http://localhost:8888/"
url:=% root . "notebooks/" . name
Run, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --profile-directory=Default --app="%url%",,OutputVarPID
return

OpenRoot:
	root:="http://localhost:8888/"
	url:=% root . "tree"
	Run, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --profile-directory=Default --app="%url%"
return

OpenFolder:
	Run, % dir
return

saveAndExit:
WinShow %processTitle%
WinClose %processTitle%
WinClose ahk_pid %OutputVarPID%
WinClose %name_no_ext% ahk_class Chrome_WidgetWin_1 ahk_exe chrome.exe
ExitApp