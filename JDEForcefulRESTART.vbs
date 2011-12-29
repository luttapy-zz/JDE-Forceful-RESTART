'
'
'*******************************************************************************
'    JDEForcefulRESTART.vbs
'    Copyright (C) 2010  EDENDEKKER.ME
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'*******************************************************************************
'
'
'============================================================================
'    __  __   __ __  __  __ __ __          __  __ _____     __ ___ 
'  ||  \|_   |_ /  \|__)/  |_ |_ /  \|    |__)|_ (_  |  /\ |__) |  
'__)|__/|__  |  \__/| \ \__|__|  \__/|__  | \ |____) | /--\| \  |  
'  			       http://blog.edendekker.me    
'
'============================================================================

If Wscript.Arguments.Count = 0 Then
	rtn = MsgBox( _
	"      __   __     __ __  __    __   __ __                  __   __  _____   __ ___ " & vbCrLf & _
	"    ||    \|_      |_  /    \|__) /     |_   |_  /    \ |       |__) |_   (_  |  /\ |__) |  " & vbCrLf & _
	"__)|__/|__    |    \__/|    \ \__ |__ |    \__/ |__   |    \ |____) | /--\|   \ |  " & vbCrLf & _
	"		   http://blog.edendekker.me     " & vbCrLf & _
	"------------------------------------------------------------------------------------" & vbCrLf & _
	" I shall now kill all JDE processes in memory if they exist. " & vbCrLf & vbCrLf & _
	" Are you sure this is what you want me to do? " & vbCrLf & _
	 "------------------------------------------------------------------------------------", _
	 4, "JDE Forceful Restart")

	If rtn = 7 Then WScript.Quit
Else
	cDeleteCache=false
	cDeleteDataItemCache=false
  For mnArgIndex = 0 to (Wscript.Arguments.Count - 1)
    If Wscript.Arguments(mnArgIndex) = "-cache" Then cDeleteCache=true
		If Wscript.Arguments(mnArgIndex) = "-dd" 	Then cDeleteDataItemCache=true
  Next
End If


'---------------------------------------------------------------------------
' SETTINGS 
'---------------------------------------------------------------------------

' These are arguments that identify JDE processes
szProcessMatchStrings	="e812"			'Process attribute matching
Const SZ_JDE_LOGIN_PATH		= "C:\e812\system\bin32\activConsole.exe"

' These are parameters that delete annoying JDE random log files
szFileMatchStrings	="jderoot*.log,jderoot*.log.lck,jas_*.log,jas_*.log.lck"
szHitPath	= "C:\Documents And Settings\%USERNAME%"
szRecursive = "/s"

'---------------------------------------------------------------------------
' MAIN 
'---------------------------------------------------------------------------
szProcessKillList		= CreateKillListByMatchString(szProcessMatchStrings)
szShellCommandTaskKill 	= CreateKillCommandsByKillList(szProcessKillList)

KillProcessByKillCommands(szShellCommandTaskKill)

WScript.Sleep 200

szDeleteList		= CreateDeleteCommandsByMatchString(szFileMatchStrings, szRecursive, szHitPath)
rtn					= DeleteFilesByDeleteCommands(szDeleteList)

WScript.Sleep 200



'Set WshShell	=	CreateObject("WScript.Shell")  'Instanciate shell handle	
'Set oExec = WshShell.Exec(SZ_JDE_LOGIN_PATH)
'Set WshShell	=	Nothing

'---------------------------------------------------------------------------
' Future Implementations
'---------------------------------------------------------------------------
' clear 6 DD Files

' clear JDE cache


'---------------------------------------------------------------------------
' Run JDE Doggy! YEAAAAHH! WOOOOOOO!
'---------------------------------------------------------------------------

Set WshShell	=	CreateObject("WScript.Shell")  'Instanciate shell handle	
Set oExec = WshShell.Exec(SZ_JDE_LOGIN_PATH)
Set WshShell	=	Nothing












'//---------------------------------------------------------------------------//
'//---------------------------------------------------------------------------//
' 							HELPER METHODS
'//---------------------------------------------------------------------------//
'//---------------------------------------------------------------------------//

'---------------------------------------------------------------------------
' INPUTS: Comma separated values of keywords as strings
' OUTPUTS: Comma separated values of ProcessIds as numbers
'---------------------------------------------------------------------------
Private Function CreateKillListByMatchString(szProcessMatchStrings)
	
	szProcessMatchStrings = Split(szProcessMatchStrings,",")
	
	' Grab all the ProcessId's where they have process attributes containing
	' JDE specific keywords
	' Add these to the process kill list! :)
	Set objService = GetObject ("winmgmts:")	
	
	' Check a set of process attributes for szMatchString values
	For Each objProcess In objService.InstancesOf ("Win32_Process")
		For each szMatchString In szProcessMatchStrings
			If InStr(UCase(objProcess.CommandLine),UCase(szMatchString))>0 OR _
				 InStr(UCase(objProcess.ExecutablePath),UCase(szMatchString))>0 OR _
				 InStr(UCase(objProcess.CommandLine),UCase(szMatchString))>0 OR _
				 InStr(UCase(objProcess.Description),UCase(szMatchString))>0 OR _
				 InStr(UCase(objProcess.Caption),UCase(szMatchString))>0Then
				' Current process was started by JDE, add to kill list
				szProcessKillList = szProcessKillList & ","	 & objProcess.ProcessId
				'MsgBox(szProcessKillList)
			End If
		Next
	
		If objProcess.Name = "iexplore.exe"	Then 
			szProcessKillList = szProcessKillList & ","	 & objProcess.ProcessId
		End If
	Next
	Set objService = Nothing
	CreateKillListByMatchString =  szProcessKillList
	
End Function

'---------------------------------------------------------------------------
' INPUTS: Comma separated values of ProcessIds as numbers
' OUTPUTS: Comma separated values of MS-DOS taskkill commands as strings
'---------------------------------------------------------------------------
Private Function CreateKillCommandsByKillList(szProcessKillList)

	szProcessKillList 	= Split(szProcessKillList,",")' convert list into an array
	
	' Iterate through the process kill list and create kill commands
	Set WshShell	=	CreateObject("WScript.Shell")  'Instanciate shell handle
	
	' Create a list of shell killing commands
	For Each szProcessId In szProcessKillList
		szShellCommandTaskKill =  szShellCommandTaskKill & _
								  "taskkill /F /PID " & _
								  szProcessId & ","
	Next
	
	Set WshShell = Nothing
	
	
	CreateKillCommandsByKillList =  szShellCommandTaskKill
	
End Function




'---------------------------------------------------------------------------
' INPUTS: Comma separated values of MS-DOS taskkill commands as strings
' OUTPUTS: None
'---------------------------------------------------------------------------
Private Function KillProcessByKillCommands(szShellCommandTaskKill)

	KillProcessByKillCommands = False
	
	szShellCommandTaskKill = Split(szShellCommandTaskKill,",")
	
	Set WshShell	=	CreateObject("WScript.Shell")  'Instanciate shell handle
	
	' Iterate through the list of kill commands and forcefully kill each process
	For Each szShellCommand In szShellCommandTaskKill
		If szShellCommand <> "" Then
			' Execute the next command in the list
			Set oExec = WshShell.Exec(szShellCommand)
			
			' Pause script execution until shell process has
			' successfully ended and is unloaded from memory space
			Do While oExec.Status = 0
				WScript.Sleep 200
			Loop		
		End If
	Next
	
	Set WshShell = Nothing
	
	KillProcessByKillCommands = True
	
End Function  

'---------------------------------------------------------------------------
' INPUTS: Comma separated values of log file match strings
' OUTPUTS: Comma separated values of MS-DOS del commands as strings
'---------------------------------------------------------------------------
Private Function CreateDeleteCommandsByMatchString(szMatchString, szRecursive, szHitPath)
	
	szMatchString = Split(szMatchString,",")
	
	For Each szItem In szMatchString
		szShellCommandDelete =  szShellCommandDelete & _
								  "cmd /c del "& szRecursive &" /q /f """ & szHitPath &"\" & szItem & ""","
	Next

	CreateDeleteCommandsByMatchString =  szShellCommandDelete
	
End Function

'---------------------------------------------------------------------------
' INPUTS: Comma separated values of MS-DOS del commands as strings
' OUTPUTS: None
'---------------------------------------------------------------------------
Private Function DeleteFilesByDeleteCommands(szShellCommandDelete)

	DeleteFilesByDeleteCommands = False
	
	szShellCommandDelete = Split(szShellCommandDelete,",")
	
	Set WshShell	=	CreateObject("WScript.Shell")  'Instanciate shell handle
	
	' Iterate through the list of kill commands and forcefully kill each process
	For Each szShellCommand In szShellCommandDelete
		If szShellCommand <> "" Then
			' Execute the next command in the list
			'MsgBox(szShellCommand)
			Set oExec = WshShell.Exec(szShellCommand)
			
			' Pause script execution until shell process has
			' successfully ended and is unloaded from memory space
			Do While oExec.Status = 0 
				WScript.Sleep 200 
			Loop		
		End If
	Next
	
	Set WshShell = Nothing
	
	DeleteFilesByDeleteCommands = True
	
End Function  



'============================================================================
'				      MADE BY http://blog.edendekker.me     
'============================================================================


