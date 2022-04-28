'******************************************************************
'* hlthchk.vbs
'*
'* This vbscript uses SQLDMO to check the status of a SQL Server(s).
'*
'* To be used after maintenance or patching, to ensure that SQL
'* is running without problems.
'* 
'* It can be used by passing in a single server name, or by passing
'* in the keyword 'file'.  If 'file' is passed in, the script will
'* read the hlthchk_svrlist.txt file for a list of servers. It will
'* do the health check on each server listed.
'*
'* Summary messages are written to standrd out, and detail is written
'* to the log file.  If the 2nd parameter is 'detail' the detail messages
'* are also written to standard out.
'*
'* 2 parameters 
'*	Parm1: required.  Either name of server or keyword 'file'
'*      Parm2: optional keyword 'detail'
'*
'* Example:
'*	cscript hlthchk.vbs servername 
'*	cscript hlthchk.vbs servername detail
'*	cscript hlthchk.vbs file 
'*
'* The script writes output messages to standard out and to a log file
'*
'*******************************************************************

Option Explicit

'************** Declarations *********************

Dim objArgs
Dim oSrvServer 
Dim oSrvDB 
Dim oSrvQR 
Dim oSrvQR2
Dim oSQLServer 
Dim oDatabase
Dim oLogin
Dim oUser
Dim oRole
Dim fso

Dim maxerr
Dim errnum
Dim errhex
Dim errdesc
Dim errmsg
Dim servernm
Dim dbname
Dim loginnm
Dim userid
Dim passwd
Dim loginty
Dim usernm
Dim rolenm
Dim prevuser
Dim newpass
Dim addrole
Dim qrymsg
Dim numresset
Dim qry

Dim totsrv
Dim totusers
Dim s
Dim x
Dim y
Dim z
Dim strlen
Dim rsrows
Dim rsscript
Dim ChangeDefdb
Dim WshShell

Dim lgfile
Dim historyfile
Dim lgfile2
Dim loginstr
Dim runrc
Dim servr2
Dim logToServer 		

Dim iIndex
Dim iCount
Dim dbStatus
Dim objDatabases
Dim numSuspect
Dim svcstatus
Dim svcRunning
Dim SearchRunning
Dim AgentRunning

Dim searchauto
Dim agentauto
Dim sqlauto

Dim svrerr

Dim namedinst
Dim svronly
Dim svr_instance
Dim strpos
Dim instanceonly

Dim oErrLog
Dim eCount
Dim eIndex
Dim numErr
Dim errPos1
Dim errline
Dim fndErr

Dim errtime
Dim errtype

Dim errlogpth
Dim fil
Dim filsiz

Dim SQLDMOLogin_NTGroup
Dim SQLDMOLogin_NTUser
Dim SQLDMOLogin_Standard
Dim ForReading

Dim getserverlist
Dim serverfile
Dim outputfile
Dim showDetail

Dim SQLDMODBStat_Normal
Dim SQLDMODBStat_Suspect 

Dim SQLDMOSvc_Running

'************ Define Constants *************

SQLDMOLogin_NTGroup = 1 
SQLDMOLogin_NTUser = 0 
SQLDMOLogin_Standard = 2 

ForReading = 1			' Open File for Reading

SQLDMODBStat_Normal = 0
SQLDMODBStat_Suspect = 256 

SQLDMOSvc_Running = 1 

'*******************************************************************
' rpad -	Pads a variable length string 
' 			with spaces on right to length of strsize
'*******************************************************************

Public Function rpad(varstr, strsize)
	Dim slen 	' String length
	on error resume next

	slen = len(varstr)		' Determine the length of the string
	if strsize > slen then 	' If necessary pad with spaces
		rpad = varstr & space(strsize - slen)
	else 
		rpad = varstr
	end if
End Function

'*******************************************************************
' lpad -	Pads a variable length string 
' 			with spaces on left to length of strsize
'*******************************************************************

Public Function lpad(varstr, strsize)
	Dim slen 	' String length
	on error resume next

	slen = len(varstr)		' Determine the length of the string
	if strsize > slen then 	' If necessary pad with spaces
		lpad = space(strsize - slen) & varstr
	else 
		lpad = varstr
	end if
End Function


'*********** Main Program Begin ************

On Error Resume Next

'***  Parse Input Parameters and setup log file ****

getserverlist = false
showDetail = false

Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
	WScript.Echo "ERROR:  Parameters not entered"
	WScript.Echo " "
	WScript.Echo "Syntax: cscript hlthchk.vbs <parm1> <parm2>"
	WScript.Echo "        parm1 (Required): server name or keyword 'file'"
	WScript.Echo "        parm2 (optional): keyword 'detail'"
	WScript.Echo ""
	WScript.Echo "        Server file is D:\Monitor\Health_check\Senthil\hlthchk_svrlist.txt"
	WScript.Echo ""
	WScript.Echo "Examples:	cscript hlthchk.vbs servername detail"
	WScript.Echo "		cscript hlthchk.vbs servername"
	WScript.Echo "		cscript hlthchk.vbs file detail"
	WScript.Echo "		cscript hlthchk.vbs file"
	WScript.Quit(1)
Else
	servernm = lcase(Trim(objArgs(0)))
	if servernm = "file" Then
		getserverlist = True
	else
		' Server name was passed in
		getserverlist = False
	
		' Check if this is a named instance of sql
		namedinst = False
		svronly = ""
		svr_instance = ""
		strpos = InStr(1,servernm,"\",1)
		If strpos > 0 Then
			'named instance
			namedinst = True
			svronly = Left(servernm, strpos-1)
			instanceonly = Right(servernm, (len(servernm)-strpos))
			svr_instance = svronly & "_" & instanceonly
		Else
			svronly = servernm
			svr_instance = servernm
		End If

	End If
	
	if objArgs.Count = 2 Then
		' check for keyword detail
		if lcase(trim(objArgs(1))) = "detail" Then
			showDetail = true
		End If
	End If	
	
	
End If

Set objArgs = Nothing

maxerr = 0

' Create log file
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

If servernm = "file" Then
	outputfile = "D:\Monitor\Health_check\Senthil\hlthchk.log"
	historyfile = "D:\Monitor\Health_check\Senthil\hlthchk_" & datepart("yyyy",now()) & "_" & datepart("m",now()) & datepart("d",now()) & "_" & datepart("h",now()) & datepart("n",now()) & datepart("s",now()) & ".log"
Else
	outputfile = "D:\Monitor\Health_check\Senthil\" & svr_instance & "_hlthchk.log"
	historyfile = "D:\Monitor\Health_check\Senthil\" & svr_instance & "_hlthchk" & "_" & datepart("yyyy",now()) & "_" & datepart("m",now()) & datepart("d",now()) & "_" & datepart("h",now()) & datepart("n",now()) & datepart("s",now()) & ".log"
End If

Err.Number = 0
Set lgfile = fso.CreateTextFile(outputfile, True)

If Err.Number <> 0 Then
	WScript.Echo "ERROR:  Creating Log File - " & outputfile
	wScript.echo "ERROR:  Script execution will still continue with no file logging..."
	wScript.echo "ERROR:  Detailed output has been automatically turned on!"
	showDetail = true
	 
	wScript.echo
	'We do not want to quit running, just stop logging to a file
	logToServer = false
else
	logToServer = true
End If

out "Hlthchk.vbs - This job started " & Now, true 


'* Call health check routine

If getserverlist = true Then

	'*****************************************
	' Open file and read list of servers
	' for each server, call healthcheck
	'*****************************************
		
	serverfile = "D:\Monitor\Health_check\Senthil\hlthchk_svrlist.txt"
	
	Err.Number = 0
	Set serverfile = fso.OpenTextFile(serverfile, ForReading)
	If Err.Number <> 0 Then
		'File does not exist.  Display error and stop
		errmsg = "ERROR:  Opening Server file: " & serverfile
		errnum = Err.Number
		errdesc = Err.Description
		maxerr = maxerr + 1
		Call errexit
	Else
		' Read each line, and call health check for each server
		x = 0
		Do While serverfile.AtEndOfStream <> True
			Err.Number = 0
			loginstr = serverfile.ReadLine
			
			If Err.Number <> 0 Then
				errmsg = "ERROR:  Reading Server File"
				errnum = Err.Number
				errdesc = Err.Description
				maxerr = maxerr + 1
				Call errexit
			Else
				' remove blanks
				loginstr = Trim(loginstr)
			
				If len(loginstr) = 0 Then
					' skip blank lines
				Else
					servernm = loginstr
					' Check if this is a named instance of sql
					namedinst = False
					svronly = ""
					svr_instance = ""
					strpos = InStr(1,servernm,"\",1)
					If strpos > 0 Then
						'named instance
						namedinst = True
						svronly = Left(servernm, strpos-1)
						instanceonly = Right(servernm, (len(servernm)-strpos))
						svr_instance = svronly & "_" & instanceonly
					Else
						svronly = servernm
						svr_instance = servernm
					End If
					Call HealthCheck	
				End If
			End If
		Loop
		serverfile.Close
	End If

Else
	Call HealthCheck

End If

out "This job finished " & Now, true

lgfile.Close

' Copy log file to history file
Set lgfile = fso.GetFile(outputfile)
lgfile.Copy(HistoryFile)


Wscript.echo
if logToServer then WScript.echo ("Details can be found in: " & outputfile) end if
Wscript.echo "Maximum Return Code is " & maxerr
WScript.Quit (maxerr)

'**********************  Main Program End *********************



'**************************************************************
'*  HealthCheck
'* 
'*  This subroutine will detemrine the status of the SQL Server
'**************************************************************
Public Sub HealthCheck()


Dim oDbaServer 
Dim oDbaDB
Dim oDbaQR
Dim totsrv
Dim y
Dim userid
Dim passwd
Dim sqlver

On Error Resume Next

' Initialize number of errors for this server
svrerr = 0

out chr(13), false
out "-----------------------------------------------", true
out "Starting health check for sql server: " & servernm, true
out chr(13), true


'***********************
'Create server object 
'***********************
Err.Number = 0
Set oSQLServer = CreateObject("SQLDMO.SQLServer2")
If Err.Number <> 0 Then
	errmsg = "ERROR:  Creating SQLDMO.SQLServer object - " & servernm
	errnum = Err.Number
	errdesc = Err.Description
	maxerr = 1
	Call errexit
End If


'************************
'Check 1 - Connect to server
'************************


oSQLServer.LoginSecure = True

Err.Number = 0
oSQLServer.Connect servernm

If Err.Number <> 0 Then
	errnum = Err.Number
	errdesc = Err.Description
	
	out "Check 1 - Connect to SQL Server:       FAIL", true
	out "Error: " & errnum, true
	out errdesc, true
	out "-----------------------------------------------", true
	out chr(13), false
	
	maxerr = maxerr + 1
	svrerr = svrerr + 1
	Exit Sub
Else
	out "-----------------------------------------------", ShowDetail
	out "Check 1 - Connect to SQL Server:       SUCCESS", true
	out "-----------------------------------------------", ShowDetail

End If

out chr(13), ShowDetail

'**********************************
'Check 2 - Status of databases
'**********************************
out "Check 2 - Status of databases" & chr(13), ShowDetail

out chr(13), false
out space(5) & rpad("Database",30) & rpad("Status",30), ShowDetail
out space(5) & rpad("---------",30) & rpad("----------",30), ShowDetail


' Find total number of databases on this server
Err.Number = 0
Set objDatabases = oSQLServer.Databases
If Err.Number <> 0 Then
	out "Error: Obtaining databases on sever " & servernm, true
	out char(13), true

	errmsg = "ERROR:  Obtaining Databases on Server - " & servernm 
	errnum = Err.Number
	errdesc = Err.Description
	Call errexit
End If
iCount = objDatabases.Count

'Check status of each database
numSuspect = 0
For iIndex = 1 to iCount

	Err.Number = 0
	dbStatus = objDatabases.Item(iIndex).Status
	If dbStatus = SQLDMODBStat_Normal Then
		out space(5) & rpad(objDatabases.Item(iIndex).Name,30) & rpad("Normal",30), showDetail
	Else 
		numSuspect = numSuspect + 1
		If dbStatus = SQLDMODBStat_Suspect Then
			out space(5) & rpad(objDatabases.Item(iIndex).Name,30) & rpad("SUSPECT",30), showDetail
		Else
			out space(5) & rpad(objDatabases.Item(iIndex).Name,30) & rpad("NOT ready for use (Status: " &  dbStatus & ")",35), showDetail
		End If
	End If	
Next

out chr(13), showDetail

out "-----------------------------------------------", showDetail
If numSuspect = 0 then
	out "Check 2 - Status of databases:         SUCCESS", true
Else
	maxerr = maxerr + 1
	svrerr = svrerr + 1
	
	out "Check 2 - Status of databases:         FAIL    ", true
End If
out "-----------------------------------------------", showDetail

out chr(13), showDetail

'**********************************
'Check 3 - Scan error log for errors
'**********************************

out "Check 3 - Scan errorlog for errors", showDetail
out chr(13), false



' Get size of errorlog for this server.

errlogpth = oSQLServer.Registry.ErrorLogPath
errlogpth = replace(errlogpth, ":", "$")
errlogpth = "\\" & svronly & "\" & errlogpth

out space(5) & "Errlog path: " & errlogpth, showDetail
out " ", showDetail

err.number = 0

Set fil = fso.GetFile(errlogpth)

If Err.Number <> 0 Then
	errmsg = "ERROR:  get file"
	errnum = Err.Number
	errdesc = Err.Description
	Call errnoexit
	Exit Sub
End If

filsiz = fil.Size
If Err.Number <> 0 Then
	errmsg = "ERROR:  get size"
	errnum = Err.Number
	errdesc = Err.Description
	Call errnoexit
	Exit Sub
End If


' If size of errorlog is 0 bytes, then do not attempt to read.
' Display message and end.
If filsiz = 0 Then

	out "** ERROR: 0 byte errorlog for server: " & servernm, true
	out "          Skipping read of errorlog", true
	out chr(13), true

	numErr = numErr + 1
	
Else

	' If file size is greater then 10 Meg, do not attempt to read it.  Display message.
	' It will hang if it tries to read it
	If filsiz > 10000000 Then
		out "** ERROR: Log file too large to read:   " & filsiz, true
		out chr(13), true
		
		numErr = numErr + 1
	Else
	
		Err.Number = 0
		Set oErrLog = oSQLServer.ReadErrorLog
		If Err.Number <> 0 Then
			maxerr = maxerr + 1	
			errmsg = "ERROR:  Reading Errorlog - " & servernm
			errnum = Err.Number
			errdesc = Err.Description
			Call errnoexit 
			Exit Sub
		End If

		numErr = 0
		fndErr = false
		eCount = oErrLog.Rows

		For eIndex = 1 To eCount
		
			sqlver =  oSQLServer.VersionMajor

			If sqlver = 8 Then
				errline = oErrLog.GetColumnString(eIndex, 1)
			else
				'sql 2005 has 3 columns to get
				errtime = oErrLog.GetColumnString(eIndex, 1)
				errtype = oErrLog.GetColumnString(eIndex, 2) 
				errline = oErrLog.GetColumnString(eIndex, 3)
			End If

			errpos1 = Instr(1, errline, "Error:", 1)
			If errpos1 > 0 Then
				If (Instr(1, errline, "SQL Network Interface library", 1) > 0) Then
					'skip this line, it's informational
				Else
					' found error
					fndErr = true
					numErr = numErr + 1

					if sqlver = 8 Then
						out space(5) & errline, showDetail
					Else
						out space(5) & errtime & " " & errtype & " " & errline, showDetail
					End If
				End If
				
			Else
				if fndErr = true then
					' display this line. It is the error text
					if sqlver = 8 Then
						out space(5) & errline, showDetail
					Else
						out space(5) & errtime & " " & errtype & " " & errline, showDetail
					End If

					fndErr = false
				End If
			End If
		Next 
	End If
End If

out chr(13), showDetail

out "-----------------------------------------------", showDetail
If numErr = 0 then
	out "Check 3 - Scan errorlog for errors:    SUCCESS", true
Else
	maxerr = maxerr + 1	
	svrerr = svrerr + 1

	out "Check 3 - Scan errorlog for errors:    FAIL", true
End If
out "-----------------------------------------------", showDetail
out chr(13), showDetail

'**********************************
'Check 4 - Verify SQLAgent and MSSearch services are running
'**********************************

out "Check 4 - Verify SQL services are running", showDetail
out chr(13), showDetail


agentRunning = true
searchRunning = true

Err.Number = 0
svcstatus = oSqlServer.JobServer.Status
If Err.Number <> 0 Then
	agentRunning = false
	
	out space(5) & "ERROR: Checking SQLServerAgent status", showDetail
	
	maxerr = maxerr + 1
	svrerr = svrerr + 1
Else
	If SvcStatus = SQLDMOSvc_Running Then
		agentRunning = true
		out space(5) & "SQLServerAgent service is running", showDetail

	Else
		agentRunning = false
		out space(5) & "ERROR: SQLServerAgent service is NOT running", showDetail

		maxerr = maxerr + 1
		svrerr = svrerr + 1
	End If
End If

' For SQL 2000 or higher, check if Microsoft Search service is running

if oSqlServer.VersionMajor = 8 Then
	' NOTE: This only works for SQL 2000.  SQL 2005 must find different way to cehck this service

	' Check if Full text is installed
	If oSqlServer.IsFullTextInstalled = true Then
		
		SvcStatus = oSqlServer.FullTextService.Status

		If Err.Number <> 0 Then
			errmsg = "ERROR:  svcstatus"
			errnum = Err.Number
			errdesc = Err.Description
			Call errnoexit
				
			SearchRunning = false
			out space(5) & "ERROR: Checking Microsoft Search status", showDetail

			maxerr = maxerr + 1	
			svrerr = svrerr + 1
		Else

			If oSqlServer.FullTextService.Status = SQLDMOSvc_Running Then
				searchRunning = true
				out space(5) & "Microsoft Search service is running", showDetail

			Else
				searchRunning = false
				out space(5) & "ERROR: Microsoft Search service is NOT running", showDetail

				maxerr = maxerr + 1	
				svrerr = svrerr + 1
			End If
		End If
	Else
		out space(5) & "Microsoft Search service is not installed", showDetail
	End If
End If

out chr(13), showDetail

out "-----------------------------------------------", showDetail
If SearchRunning = true and AgentRunning = true then
	out "Check 4 - Verify Services are running: SUCCESS", true
Else
	out "Check 4 - Verify Services are running: Fail", true
End If
out "-----------------------------------------------", showDetail
out chr(13), showDetail


'*****************************************************************
'Check 5 - Verify SQL Server services are set to Auto start
'******************************************************************

out "Check 5 - Verify SQL services are set to Auto start", showDetail
out chr(13), showDetail

agentauto = true
sqlauto = true

' If SQL is clustered, services must be set to Manual.  Skip this check
if oSQLServer.IsClustered = true then
	out space(5) & "SQL Server is clustered.  Services must be set to Manual start", showDetail

Else

	' For SQL 2000 or higher, check MSSQLServer start up
	if oSqlServer.VersionMajor >= 8 Then

		Err.Number = 0
		svcstatus = oSqlServer.AutoStart
		If Err.Number <> 0 Then
			sqlauto = false
			
			out space(5) & "ERROR: Checking SQL Server autostart", showDetail

			maxerr = maxerr + 1
			svrerr = svrerr + 1
		Else
			If SvcStatus = True Then
				sqlauto = true
				out space(5) & "SQL Server service is set to Auto Start", showDetail
			Else
				sqlauto = false
				out space(5) & "ERROR: SQL Server service is set to Manual start", showDetail

				maxerr = maxerr + 1
				svrerr = svrerr + 1
			End If
		End If

	End If

	' Check SQL Agent start up
	Err.Number = 0
	svcstatus = oSqlServer.JobServer.AutoStart
	If Err.Number <> 0 Then
		agentauto = false
		out space(5) & "ERROR: Checking SQLServerAgent autostart", showDetail

		maxerr = maxerr + 1
		svrerr = svrerr + 1
	Else
		If SvcStatus = True Then
			agentauto = true
			out space(5) & "SQLServerAgent service is set to Auto Start", showDetail

		Else
			agentauto = false
			out space(5) & "ERROR: SQLServerAgent service is set to Manual start", showDetail

			maxerr = maxerr + 1
			svrerr = svrerr + 1
		End If
	End If
End If

out chr(13), showDetail

out "-----------------------------------------------", showDetail
If sqlauto = true and Agentauto = true then
	out "Check 5 - Verify Services startup:     SUCCESS", true
Else
	out "Check 5 - Verify Services startup:     FAIL", true
End If
out "-----------------------------------------------", showDetail





'**************************
'*  Disconnect From Server 
'**************************
oSQLServer.Disconnect
set oSQLServer = nothing

out "-----------------------------------------------", showDetail
out "Completed health check for sql server: " & servernm, showDetail
out "Total errors for server:               "  & svrerr, true
if svrerr = 0 Then
	out "Status:                                SUCCESS", true
Else
	out "Status:                                FAIL", true
End If
out "-----------------------------------------------", true
out chr(13), true


End Sub


public function out(output, echoMe)

	if logToServer then 
		lgFile.writeline(output)
	end if
	
	if echoMe then
		wscript.echo(output)
	end if
	
	out = true
	
end function



'************  errexit Subroutine ***********************************
' This section will display error information and exit the program

Public Sub errexit()

on error resume next

	Wscript.echo errmsg
	Wscript.echo "ErrNumber: " & errnum & " " & errhex
	Wscript.echo errdesc
	Wscript.echo " "
	Wscript.echo "Maximum Return Code is " & maxerr
	Wscript.echo " "
	
	WScript.Quit (maxerr)
End Sub

'************  errnoexit Subroutine ***********************************
' This section will display error information but NOT exit the program

Public Sub errnoexit()

on error resume next

	Wscript.echo errmsg
	Wscript.echo "ErrNumber: " & errnum & " " & errhex
	Wscript.echo errdesc
		
End Sub