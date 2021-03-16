'************************************************************
'* DESCRIPTION: Folder Monitoring for SFTP Servers.
'*              
'* 
'* AUTHOR: Arnold Sibuea
'*
'* FUNCTION :
'* 1. Monitor Files on SFTP Folders.
'* 2. Move file from SFTP Folder to Archive folder if the Created Date is more than 2 Weeks
'* 3. Delete File from Archive Folder if Created date is more than 4 Weeks
'* 4. Create Log files for the ScriptEngine
'* 5. Run the script on multiple Folders
'* 6. Add Error Handling
'************************************************************
On Error Resume Next
Err.Clear
Const ForReading = 1
Const ADS_PROPERTY_APPEND = 3

'Parameter
moveDaysOld = 5     						    					 'Days Before moving to Archive Folder
delDaysOld = 10						    						 'Days Before File is deleted
waitTime = 2000	   					    							 'Wait Time Before Checking the Next SFTP Folder
WaitFolder = 86400000 					    						 'Wait Time Before Checking the Archive Folder
strLogFileName = "C:\Script\Monitoring\Monitoringlog.txt"       'Log Files Name
Today = Date

'Check if Script already running.
'Quit if it does

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
("Select * from Win32_Process Where Name = 'cscript.exe'")

' Not Used, only for reference if you want to check the cscript also
' ("Select * from Win32_Process Where Name = 'cscript.exe'" & _
' " OR Name = 'wscript.exe'")

script2 = str(WScript.ScriptName)
intCount = 0
For Each objItem in colItems
    x= Replace(objItem.CommandLine, "cscript  ","")
    If Script1 = script2 Then
    	intCount = Int(intCount) + 1
    End If
Next

If IntCount > 1 Then
Wscript.Quit
End If

' Open List.csv file that contain all listed folder that need to be monitored

strListName = "C:\Script\Monitoring\FolderList.csv"

'Connect to FSO object
Set FSO = CreateObject("Scripting.FileSystemObject")
set objShell = WScript.CreateObject ("WScript.Shell")
strDirectory = objShell.CurrentDirectory

'Check if the Directory folder File List is exsist.
if fso.FileExists(strListName) then
      Call WriteLogFileLine (strLogFileName, "Directory List file is found")
Else
      Call WriteLogFileLine (strLogFileName, "Directory List file is not found!!")
      Call WriteLogFileLine (strLogFileName, "Program is now quiting !!")
      wscript.Quit	  
End If

'  Check if Log files is exsist
'  if not Create the log files

if fso.FileExists(strLogFileName) then
      Call WriteLogFileLine (strLogFileName, "Log file is found")
Else
      Call WriteLogFileLine (strLogFileName, "directory List file is not found!!")
      Call WriteLogFileLine (strLogFileName, "Creating the log files.")
      Set objlog = FSO.CreateTextFile(strDirectory & strFile)
      Call WriteLogFileLine (strLogFileName, "log files is created")
        
End If

' Add Text to the log files to acknowledge the script is Running

new1 = "*************************************************"
new2 = "Starting to check the folders on " & Now() 

Call WriteLogFileLine(strLogFileName,new1)
Call WriteLogFileLine(strLogFileName,new2)
Call WriteLogFileLine(strLogFileName,new1)

	  
Set objTextFile = FSO.OpenTextFile(strListName, ForReading)

Do Until objTextFile.AtEndOfStream
strFolder = objTextFile.ReadLine

  Set folder = fso.getFolder(strFolder & "\SFTP\")
  set ArcFolder = fso.getFolder(strFolder & "\Archive\")

  Set colFiles = folder.Files
  
  CheckFolder = "Checking " & strFolder & "\SFTP"
  Call WriteLogFileLine(strLogFileName,CheckFolder)

  if folder.files.Count = 0 then
    Call WriteLogFileLine(strLogFileName,"No Files Found on " & strFolder & "\SFTP")
  Else
    Call WriteLogFileLine(strLogFileName, folder.Files.count & " Files Found on " & strFolder & "\SFTP")
  End If

  Err.Clear
  For Each objFile in colFiles
		If Err.Number <> 0 Then
			WriteErrX1 =  "Error: " & Err.Number
			WriteErrX2 =  "Error (Hex): " & Hex(Err.Number)
			WriteErrX3 =  "Source: " &  Err.Source
			WriteErrX4 =  "Description: " &  Err.Description
			WriteErrX5 =  "Line: " & err.Line
			Call WriteLogFileLine(strLogFileName,"Write X : " & WriteErrX1)
			Call WriteLogFileLine(strLogFileName,WriteErrX2)
			Call WriteLogFileLine(strLogFileName,WriteErrX3)
			Call WriteLogFileLine(strLogFileName,WriteErrX4)
			Call WriteLogFileLine(strLogFileName,WriteErrX5)
			Err.Clear
		 End If
		 
      FileCheck =  objFile & " Created on " & objFile.DateCreated & "."
      Call WriteLogFileLine(strLogFileName,FileCheck)
	  
	  objCreateDate = cstr(objFile.DateCreated)
	  'Call WriteLogFileLine(strLogFileName,objCreateDate)
	  
	  intAge = datediff("d",objFile.DateCreated,Today)
	  
	  objDaytoMove = "File has been stored for " & Int(intAge) & " Days"
          Call WriteLogFileLine(strLogFileName,objDaytoMove)
	  Call WriteLogFileLine(strLogFileName,"File will be moved to Archive Folder in " & Cstr(moveDaysOld-int(intAge)) & " Days")
      
      if int(intAge) >= moveDaysOld Then
	  
		'Write to Log first
		 objFileMove = "Moving File  " & objFile & "."
	     Call WriteLogFileLine(strLogFileName,objFileMove)
	     
		 'Move the File
	     FSO.MoveFile cstr(objFile), strFolder & "\Archive\"
		 If Err.Number <> 0 Then
			WriteErr11 =  "Error: " & Err.Number
			WriteErr12 =  "Error (Hex): " & Hex(Err.Number)
			WriteErr13 =  "Source: " &  Err.Source
			WriteErr14 =  "Description: " &  Err.Description
			WriteErr15 =  "Line: " & err.Line
			Call WriteLogFileLine(strLogFileName,"Write 1X : " & WriteErr11)
			Call WriteLogFileLine(strLogFileName,WriteErr12)
			Call WriteLogFileLine(strLogFileName,WriteErr13)
			Call WriteLogFileLine(strLogFileName,WriteErr14)
			Call WriteLogFileLine(strLogFileName,WriteErr15)
			Err.Clear
		 End If
      End If
Next

Border = "------------------------------------------------------------------------"
Call WriteLogFileLine(strLogFileName,Border)
'Wscript.Sleep waitTime


Set arrFiles = Arcfolder.Files
  

  CheckFolder2 = "Checking " & strFolder & "\Archive"
  Call WriteLogFileLine(strLogFileName,CheckFolder2)
  
  if ArcFolder.files.Count = 0 then
    Call WriteLogFileLine(strLogFileName,"No Files Found on " & strFolder & "\Archive")
  Else
    Call WriteLogFileLine(strLogFileName, ArcFolder.Files.count & " Files Found on " & strFolder & "\Archive")
  End If

  For Each objFile2 in arrFiles

       FileCheck3 =  objFile2 & " Created on " &  objFile2.DateCreated & "."
       Call WriteLogFileLine(strLogFileName,FileCheck3)
       intAge2 = DateDiff("d",objFile2.DateCreated,Today)
       objDaytoDel = "File has been stored for : " & int(intAge2) & " Days"
       Call WriteLogFileLine(strLogFileName,objDaytoDel)
       Call WriteLogFileLine(strLogFileName,"File will be deleted from archive folfer in " & Cstr(delDaysOld-int(intAge2)) & " Days")
       
	   if Int(intAge2) >= delDaysOld Then
	  
         ' Write to logs
         objFileDelete = "Deleting File : " & objFile2
	     Call WriteLogFileLine(strLogFileName,objFileDelete)
                   
   	 ' Delete the files
	     Err.Clear
         fso.DeleteFile  cstr(objFile2)
		 If Err.Number <> 0 Then
			WriteErr21 =  "Error: " & Err.Number
			WriteErr22 =  "Error (Hex): " & Hex(Err.Number)
			WriteErr23 =  "Source: " &  Err.Source
			WriteErr24 =  "Description: " &  Err.Description
			WriteErr25 =  "Line: " & Err.Line
			Call WriteLogFileLine(strLogFileName,WriteErr21)
			Call WriteLogFileLine(strLogFileName,WriteErr22)
			Call WriteLogFileLine(strLogFileName,WriteErr23)
			Call WriteLogFileLine(strLogFileName,WriteErr24)
			Call WriteLogFileLine(strLogFileName,WriteErr25)
			Err.Clear
		 End If

       End If
   Next
Call WriteLogFileLine(strLogFileName,Border)   
Loop

objTextFile.close

' add the ending for every check

last1 = "*************************************************"
last2 = "End of Folder Checking on " & Now()

Call WriteLogFileLine(strLogFileName,last1)
Call WriteLogFileLine(strLogFileName,last2)
Call WriteLogFileLine(strLogFileName,last1)
Call WriteLogFileLine(strLogFileName,"")
Call WriteLogFileLine(strLogFileName,"")
Call WriteLogFileLine(strLogFileName,"")


'Wscript.Sleep WaitFolder




'function to Add data to the log files
Function WriteLogFileLine(sLogFileName,sLogFileLine)
    dateStamp = Now()

    Set objFsoLog = CreateObject("Scripting.FileSystemObject")
    Set logOutput = objFsoLog.OpenTextFile(sLogFileName, 8, True)

    logOutput.WriteLine(cstr(dateStamp) + " -" + vbTab + sLogFileLine)
    logOutput.Close

    Set logOutput = Nothing
    Set objFsoLog = Nothing
End Function

Function stripchars(s1,s2)
	For i = 1 To Len(s1)
		If InStr(s2,Mid(s1,i,1)) Then
			s1 = Replace(s1,Mid(s1,i,1),"")
		End If
	Next
	stripchars = s1
End Function