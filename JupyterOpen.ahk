#Persistent
#SingleInstance off

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
	if(GetKeyState("Shift", "P") && ext=="ipynb")
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
		
		RunWait, jupyter nbconvert "%name%" --to pdf --template article
		
		if(!FileExist(file))
			msgbox, Error creating the file!
		else
			SelectFile(file)
		GoSub, saveAndExit
	}
	IniRead, root, %A_ScriptDir%\JupyterOpen.ini, %dir%,root,0
	IniRead, token, %A_ScriptDir%\JupyterOpen.ini, %dir%,token,0
	if(root=0)
	{
		url:=StdoutToVar_CreateProcess("jupyter notebook --no-browser", "http.*$")
		RegExMatch(url, "http.*localhost.*?/", root)
		RegExMatch(url, "token\=.*", token)
		IniWrite, %root%, %A_ScriptDir%\JupyterOpen.ini, %dir%,root
		IniWrite, %token%, %A_ScriptDir%\JupyterOpen.ini, %dir%,token
	}
	if(ext=="ipynb")
		url:=% root . "notebooks/" . name . "?" . token
	else
		url:=% root . "tree" . "?" . token
	Run, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --profile-directory=Default --app="%url%"
	if(!iProcessId){
		GoSub, saveAndExit
	}
	else
	{
		Menu, Tray, Tip, %dir%
		Menu, Tray, NoStandard
		Menu, Tray, Add, &Exit, saveAndExit
		Menu, Tray, add, &Open Browser, OpenRoot
		Menu, Tray, add, &Open Folder, OpenFolder
	}
}

return

OpenRoot:
	;Run, % root . "tree"
	url:=% root . "tree"
	Run, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --profile-directory=Default --app="%url%"
return

OpenFolder:
	Run, % dir
return

saveAndExit:
if(iProcessId)
{
	KillChildProcesses(iProcessId)
	Process, Close, %iProcessId%
	IniDelete, %A_ScriptDir%\JupyterOpen.ini, %dir%,root
	IniDelete, %A_ScriptDir%\JupyterOpen.ini, %dir%,token
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