#Persistent
#SingleInstance off

shiftPressed:=GetKeyState("Shift", "P")
controlPressed:=GetKeyState("Control", "P")

jupyter_nbconvert_path:="""C:\Users\langh\Miniconda3\Scripts\jupyter-nbconvert.exe"""
jupyter_notebook_path:="""C:\Users\langh\Miniconda3\Scripts\jupyter-notebook.exe""  --no-browser"
;command="C:\Windows\System32\bash.exe" -ilc "source activate xeus; jupyter notebook"
;jupyter_notebook_path=%comspec% /k "%command%"

if(GetProcessCount()=1)
	FileDelete, %A_ScriptDir%\JupyterOpen.ini

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
SplitPath, args, name, dir, ext, name_no_ext, drive
FileGetAttrib, Attributes, %args%		
if(InStr(Attributes, "D"))
{
	dir:=args
}
SetWorkingDir, %dir%
OnExit, saveAndExit
if(ext=="ipynb" || InStr(Attributes, "D"))
{
	if(shiftPressed && ext=="ipynb")
	{
		export:=1
		file:=dir . "\" . name_no_ext . ".pdf"
		if(FileExist(file))
			export:=0
		if(!export)
		{
			MsgBox, 4,Overwrite file!?, Overwrite file? (Si o No)
			IfMsgBox Yes
			{
				export:=1
			}
		}
		
		if(!export)
			GoSub, saveAndExit
		
		RunWait, %jupyter_nbconvert_path% "%dir%/%name%" --to pdf --template "better-article.tplx", %A_ScriptDir%
		if(!FileExist(file))
			msgbox, Error creating the file!
		else
			SelectFile(file)
		GoSub, saveAndExit
	}
	IniRead, root, %A_ScriptDir%\JupyterOpen.ini, %dir%,root,0
	IniRead, token, %A_ScriptDir%\JupyterOpen.ini, %dir%,token,0
	if(root=0 || root="")
	{
		url:=StdoutToVar_CreateProcess(jupyter_notebook_path, "http.*$")
		RegExMatch(url, "http.*localhost.*?/", root)
		RegExMatch(url, "token\=.*", token)
		IniWrite, %root%, %A_ScriptDir%\JupyterOpen.ini, %dir%,root
		IniWrite, %token%, %A_ScriptDir%\JupyterOpen.ini, %dir%,token
	}
	if(ext=="ipynb")
	{
		url:=% root . "notebooks/" . name . "?" . token
		IniRead, opened, %A_ScriptDir%\JupyterOpen.ini, %dir%, opened, 0
		openedStr:="=" . opened
		if (!InStr(openedStr, "|" . name) && !InStr(openedStr, "=" . name)){
			opened:=name . (opened<>"0" ? "|" . opened : "")
			IniWrite, %opened%, %A_ScriptDir%\JupyterOpen.ini, %dir%, opened
		}
	}
	else
		url:=% root . "tree" . "?" . token
	Run, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --profile-directory=Default --app="%url%"
	if(!iProcessId){
		GoSub, saveAndExit
	}
	else
	{
		SplitPath, dir , name2, dir2, ext2, name_no_ext2, drive2
		Menu, Tray, Tip, %dir%
		Menu, Tray, NoStandard
		Menu, Tray, Add, &Exit, saveAndExit
		Menu, Tray, add, &Open Browser, OpenRoot
		Menu, Tray, add, &Open Folder, OpenFolder
		Menu, Tray, Add
		Menu, Tray, Add, %name2%, OpenFolder
		Menu, Tray, Add
		SetTimer, BuildMenu, 500
	}
}
SetTimer, CheckAlive, 10000
return

BuildMenu:
	IniRead, opened, %A_ScriptDir%\JupyterOpen.ini, %dir%, opened, 0
	if(opened<>"0" && opened<>oldOpenedMenu)
	{
		if(oldOpenedMenu)
		{
			count:=1
			Loop,Parse,oldOpenedMenu,|
			{
				Menu, Tray, delete, & %count% %A_LoopField%
				count:=count+1
			}
		}
		count:=1
		Loop,Parse,opened,|
		{
			Menu, Tray, add, & %count% %A_LoopField%, OpenRecentFile
			count:=count+1
		}
		oldOpenedMenu:=opened
	}
return

OpenRecentFile:
	recentFile:=LTrim(RTrim(SubStr(A_ThisMenuItem, 4)))
	recentUrl:=% root . "notebooks/" . recentFile . "?" . token
	Run, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --profile-directory=Default --app="%recentUrl%"
return

OpenRoot:
	;Run, % root . "tree"
	url:=% root . "tree"
	Run, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --profile-directory=Default --app="%url%"
return

OpenFolder:
	Run, % dir
return

saveAndExit:
SetTimer, BuildMenu, Off
if(iProcessId)
{
	KillChildProcesses(iProcessId)
	Process, Close, %iProcessId%
	IniDelete, %A_ScriptDir%\JupyterOpen.ini, %dir%,root
	IniDelete, %A_ScriptDir%\JupyterOpen.ini, %dir%,token
	IniDelete, %A_ScriptDir%\JupyterOpen.ini, %dir%,opened
	if(GetProcessCount()=1)
		FileDelete, %A_ScriptDir%\JupyterOpen.ini
}
ExitApp
	
StdoutToVar_CreateProcess(sCmd, searchString="", sEncoding:="CP0", sDir:="", ByRef nExitCode:=0) {
	global iProcessId
    DllCall( "CreatePipe",           PtrP,hStdOutRd, PtrP,hStdOutWr, Ptr,0, UInt,0 )
    DllCall( "SetHandleInformation", Ptr,hStdOutWr, UInt,1, UInt,1                 )

            VarSetCapacity( pi, (A_PtrSize == 4) ? 16 : 24,  0 )
    siSz := VarSetCapacity( si, (A_PtrSize == 4) ? 68 : 104, 0 )
    NumPut( siSz,      si,  0,                          "UInt" )
    NumPut( 0x100,     si,  (A_PtrSize == 4) ? 44 : 60, "UInt" )
    NumPut( hStdOutWr, si,  (A_PtrSize == 4) ? 60 : 88, "Ptr"  )
    NumPut( hStdOutWr, si,  (A_PtrSize == 4) ? 64 : 96, "Ptr"  )

    If ( !DllCall( "CreateProcess", Ptr,0, Ptr,&sCmd, Ptr,0, Ptr,0, Int,True, UInt,0x08000000
                                  , Ptr,0, Ptr,sDir?&sDir:0, Ptr,&si, Ptr,&pi ) )
        Return ""
      , DllCall( "CloseHandle", Ptr,hStdOutWr )
      , DllCall( "CloseHandle", Ptr,hStdOutRd )
	iProcessId := NumGet(pi, 16, "UInt")
    DllCall( "CloseHandle", Ptr,hStdOutWr ) ; The write pipe must be closed before reading the stdout.
    While ( 1 )
    { ; Before reading, we check if the pipe has been written to, so we avoid freezings.
        If ( !DllCall( "PeekNamedPipe", Ptr,hStdOutRd, Ptr,0, UInt,0, Ptr,0, UIntP,nTot, Ptr,0 ) )
            Break
        If ( !nTot )
        { ; If the pipe buffer is empty, sleep and continue checking.
            Sleep, 100
            Continue
        } ; Pipe buffer is not empty, so we can read it.
        VarSetCapacity(sTemp, nTot+1)
        DllCall( "ReadFile", Ptr,hStdOutRd, Ptr,&sTemp, UInt,nTot, PtrP,nSize, Ptr,0 )
        sOutput .= StrGet(&sTemp, nSize, sEncoding)
		if(searchString<>"" && RegExMatch(sOutput, searchString, match))
		{
			sOutput:=match
			break
		}
    }
    
    ; * SKAN has managed the exit code through SetLastError.
    DllCall( "GetExitCodeProcess", Ptr,NumGet(pi,0), UIntP,nExitCode )
    DllCall( "CloseHandle",        Ptr,NumGet(pi,0)                  )
    DllCall( "CloseHandle",        Ptr,NumGet(pi,A_PtrSize)          )
    DllCall( "CloseHandle",        Ptr,hStdOutRd                     )
    Return sOutput
}

KillChildProcesses(ParentPidOrExe){
	static Processes, i
	ParentPID:=","
	If !(Processes)
		Processes:=ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
	i++
	for Process in Processes
		If (Process.Name=ParentPidOrExe || Process.ProcessID=ParentPidOrExe)
			ParentPID.=process.ProcessID ","
	for Process in Processes
		If InStr(ParentPID,"," Process.ParentProcessId ","){
			KillChildProcesses(process.ProcessID)
			Process,Close,% process.ProcessID 
		}
	i--
	If !i
		Processes=
}

SelectFile(file)
{
	StringReplace, file, file,/,\, All
	char=`" ;"
	while(SubStr(file, 1,1)=char)
		file:=Trim(SubStr(file,2,StrLen(file)-1))
	while(SubStr(file, StrLen(file),1)=char)
		file:=Trim(SubStr(file,1,StrLen(file)-1))
	char=`\
	while(SubStr(file, StrLen(file),1)=char)
		file:=Trim(SubStr(file,1,StrLen(file)-1))
	SplitPath, file, name, dir, ext, name_no_ext, drive
	folder:=dir
	wins := ComObjCreate("Shell.Application").windows
	Run, "%folder%"
	if(FileExist(file)){ ;name<>name_no_ext && 
		buscarVentana:=1
		while(buscarVentana){
			For win in wins
			{
				 ComObjError(false)
				if(win.document.folder)
				if(win.document.folder.self.path=folder)
				{
					doc:=win.document
					buscarVentana:=0
				}
			}		
		}
		ComObjError(true)
		items := doc.folder.items
		For item in items
		{		
			;if(item.path=file)
			;	doc.SelectItem(item, 15)
			;else
				doc.SelectItem(item, 16)
		}
		doc.SelectItem(items.item(name), 16)	
		doc.SelectItem(items.item(name), 1)
	}
}

GetProcessCount(){
	count=0
	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
	{
		if process.Name = "JupyterOpen.exe"
			count++
	}
	return count
}

CheckAlive:
if(!chromeInstance1)
	chromeInstance1:=WinExist("Home ahk_exe chrome.exe")
if(!chromeInstance2)
	chromeInstance2:=WinExist("Jupyter Notebook ahk_exe chrome.exe")
if(!chromeInstance3)
	chromeInstance3:=WinExist(chromeTitle . " ahk_exe chrome.exe")
IniRead, opened, %A_ScriptDir%\JupyterOpen.ini, %dir%, opened, 0
count:=0
Loop,Parse,opened,|
{
	DetectHiddenWindows, On
	chromeTitle:=StrReplace(A_LoopField, ".ipynb" , "")
	if(WinExist("ahk_id " . chromeInstance1) || WinExist("ahk_id " . chromeInstance2) || WinExist("ahk_id " . chromeInstance3))
		count:=count+1
}
if(!count)
	GoSub, saveAndExit
return
