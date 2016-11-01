rem Install/Update a Windows Application using VBS
rem Diego Costa Bürger
rem 2008-09-23

Dim oShell
Dim oProcEnv
Dim fso
Dim sProgramFiles

Set oShell = Wscript.CreateObject("Wscript.Shell")
Set oProcEnv = oShell.Environment("PROCESS")
Set fso = CreateObject("Scripting.FileSystemObject")
sProgramFiles = oProcEnv("ProgramFiles")

Call CreateRunInStart()
Call CreateStartMenu()
Call DownloadAndRun()

resp = MsgBox("Successfully Install/Update!", 0, "SoftwareName")

Sub CreateRunInStart
	On Error Resume Next

	oShell.RegDelete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\SoftwareName"
	oShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\SoftwareName"

	If sProgramFiles = "" Then
	  sProgramFiles = oShell.RegRead _
		 ("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir")
	End If

	oShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\SoftwareName", sProgramFiles&"\SoftwareName\software.exe", "REG_SZ"
	Set oShell = Nothing
End Sub

Sub CreateStartMenu
	On Error Resume Next

	Const START_MENU = &Hb&
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.Namespace(START_MENU) 
	Set objFolderItem = objFolder.Self 
	If Not fso.FolderExists(objFolderItem.Path&"\Programas") Then
		sFolder = objFolderItem.Path&"\Programs\SoftwareName"
	Else
		sFolder = objFolderItem.Path&"\Programas\SoftwareName"
	End If
	Set newfolder = fso.CreateFolder(sFolder)

	Set objWSHShell = CreateObject("WScript.Shell")
	Set objSC = objWSHShell.CreateShortcut(sFolder&"\software_name.lnk") 
	objSC.Description = "Send"
	objSC.HotKey = "CTRL+ALT+SHIFT+X"
	objSC.IconLocation = "software.exe, 0"  ' 0 is the index
	objSC.TargetPath = sProgramFiles&"\SoftwareName\software.exe"
	objSC.WindowStyle = 1   ' 1 = normal; 3 = maximize window; 7 = minimize
	objSC.WorkingDirectory = sProgramFiles&"\SoftwareName\"
	objSC.Save

End Sub

Sub DownloadAndRun
	On Error Resume Next

	strFileURL = "http://www.myremotehost.com/software_name/SoftwareName.zip"
	strHDLocation = "C:\SoftwareName.zip"
	fso.DeleteFile Replace(strHDLocation, "\", "\\"), True

	If Not fso.FolderExists(sProgramFiles&"\SoftwareName") Then
		Set newfolder = fso.CreateFolder(sProgramFiles&"\SoftwareName")
	End If

	fso.DeleteFolder sProgramFiles&"\SoftwareName\tcl", True
	fso.DeleteFolder sProgramFiles&"\SoftwareName\tcl8.4", True
	fso.DeleteFolder sProgramFiles&"\SoftwareName\tk8.4", True

	DeleteFilesa fso.GetFolder(sProgramFiles&"\SoftwareName"), "exe"
	Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")

	objXMLHTTP.open "GET", strFileURL, false
	objXMLHTTP.send()
	 
	If objXMLHTTP.Status = 200 Then
	  Set objADOStream = CreateObject("ADODB.Stream")
	  objADOStream.Open
	  objADOStream.Type = 1 'adTypeBinary

	  objADOStream.Write objXMLHTTP.ResponseBody
	  objADOStream.Position = 0

	  Set objFSO = Createobject("Scripting.FileSystemObject")
		If objFSO.Fileexists(Replace(strHDLocation, "\", "\\")) Then objFSO.DeleteFile Replace(strHDLocation, "\", "\\")
	  Set objFSO = Nothing

	  objADOStream.SaveToFile strHDLocation
	  objADOStream.Close
	  Set objADOStream = Nothing
	  Extract strHDLocation, sProgramFiles&"\SoftwareName\"
	End if
	Set objXMLHTTP = Nothing

End Sub

Sub DeleteFilesa(srcFolder, strExt)
	Dim srcFile
	If srcFolder.Files.Count = 0 Then
		Exit Sub
	End If
	For Each srcFile in srcFolder.Files
		'MsgBox CStr(srcFile),65,""
		'If LCase(Right(srcFile.Name, Len(strExt))) = strExt Then 'Não verifica extensão
		   'If DateDiff("d", Now, srcFile.DateLastModified) < -2 Then
			   fso.DeleteFile srcFile, True
		   'End If
		'End If
	Next

End Sub


Sub Extract( ByVal myZipFile, ByVal myTargetDir )

	Dim intOptions, objShell, objSource, objTarget
    Set objShell = CreateObject( "Shell.Application" )
    Set objSource = objShell.NameSpace( myZipFile ).Items( )
    Set objTarget = objShell.NameSpace( myTargetDir )

    ' These are the available CopyHere options, according to MSDN
    ' (http://msdn2.microsoft.com/en-us/library/ms723207.aspx).
    ' On my test systems, however, the options were completely ignored.
    '      4: Do not display a progress dialog box.
    '      8: Give the file a new name in a move, copy, or rename
    '         operation if a file with the target name already exists.
    '     16: Click "Yes to All" in any dialog box that is displayed.
    '     64: Preserve undo information, if possible.
    '    128: Perform the operation on files only if a wildcard file
    '         name (*.*) is specified.
    '    256: Display a progress dialog box but do not show the file
    '         names.
    '    512: Do not confirm the creation of a new directory if the
    '         operation requires one to be created.
    '   1024: Do not display a user interface if an error occurs.
    '   4096: Only operate in the local directory.
    '         Don't operate recursively into subdirectories.
    '   9182: Do not copy connected files as a group.
    '         Only copy the specified files.
    intOptions = 16

    ' UnZIP the files
    objTarget.CopyHere objSource, intOptions

    ' Release the objects
    Set objSource = Nothing
    Set objTarget = Nothing
    Set objShell  = Nothing
End Sub 