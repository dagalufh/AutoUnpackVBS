' Created by Mikael Aspehed
' Version 0.1

' Define the global objects needed
Dim objShell, fso, objShell_wscript, FoundDirectories, ValidExtensions
Set objShell = CreateObject ("Shell.Application")
Set objShell_wscript= CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

If (Wscript.Arguments.Count < 1) Then
	MsgBox "At least one parameter required. Usage: AutoUnpack.vbs PathToSource [DeleteAfter]"
	Wscript.Quit
End If
SourceRoot = WScript.Arguments.Item(0)
SecondArgument = False


If (Wscript.Arguments.Count > 1) Then
	If (WScript.Arguments.Item(1) = DeleteAfter) Then
		SecondArgument = True
	End If
End If

FoundDirectories = array()

DirectoryFolder = fso.GetParentFolderName(WScript.ScriptFullName)
Executable = DirectoryFolder & "\Binary\unrar.exe "
PathToExtraction =  chr(34) & Executable & chr(34) & "x -o+ -y %TARGETFILE% %TARGETPATH%" 


If Not fso.FileExists(Executable) Then
	Msgbox "Unable to find Unrar.exe, it should be located: " & vbCrlf & Executable
	WScript.Quit
End If

Sub SourceDirectory (path, action)
	Set objFolder = objShell.Namespace(path)
	if objFolder is nothing Then
		MsgBox "Error!! " & path
	End If
	
	For Each strFileName In objFolder.Items
		
		Found = False
		
		' Check if it is a folder, if so, call the SourceDirectory again to search that subfolder.
		if fso.FolderExists(path & "\" & strFileName) then
			SourceDirectory  path & "\" & strFileName, action
		else
			
			'FileExtension = Right(strFileName,Len(strFileName)-inStrRev(strFileName,"."))
			FileExtension = Split(strFileName,".")
			'MsgBox FileExtension(Ubound(FileExtension)-1) &  "." & FileExtension(Ubound(FileExtension))
			
			If (Action = "Unpack") Then
				
				if (instr(FileExtension(Ubound(FileExtension)-1),"part01")) and (FileExtension(Ubound(FileExtension)) = "rar") Then
					'Msgbox "A Part01.rar file. This can be unpacked." & "(" & strFileName & ")"
					Found = True
				End If
				
				if Not (instr(FileExtension(Ubound(FileExtension)-1),"part")) and (FileExtension(Ubound(FileExtension)) = "rar") Then
					'Msgbox "A file.rar file. This can be unpacked." & "(" & strFileName & ")"
					Found = True
				End If
				
				If Found Then
					ReDim Preserve FoundDirectories(UBound(FoundDirectories)+1)
					FoundDirectories(UBound(FoundDirectories)) = path
					PathToExtraction = Replace(PathToExtraction,"%TARGETPATH%", chr(34) & path & "\" & chr(34))
					PathToExtraction = Replace(PathToExtraction,"%TARGETFILE%", chr(34) & path & "\" & strFileName &  chr(34))
					
					'MsgBox "cmd /c " & PathToExtraction
					'objShell_wscript.LogEvent 4,  PathToExtraction
					objShell_wscript.LogEvent 4,  "Attempting to execute: " & PathToExtraction & " Another logevent will be posted after either successfull or failure."
					Result = objShell_wscript.Run(PathToExtraction,0,True)
					If Result <> 0 Then
						objShell_wscript.LogEvent 4, "Something went wrong. Unrar ended with result: " & Result & " when the following command was executed: " &  PathToExtraction
					Else
						objShell_wscript.LogEvent 4, "Successfully extracted from " & strFileName
					End If
				End If
				
			Else
			' Delete them!
				if (InStr(FileExtension(Ubound(FileExtension)),"r") = 1) Then
					'Msgbox "Deleted a rar file: " & strFileName
					objShell_wscript.LogEvent 4, "Deleted file " & strFileName & " ( Not yet implemented, this is currently for information. ) "
				End If
			End If
			
		End if				
	Next
End Sub

' Start the search and unpacking...
SourceDirectory "H:\Storage\Games\Shift.2.Unleashed-RELOADED", "Unpack"

' If user called this script with a second parameter, DeleteAfter, then we go through this and deletes all *.r* files.
If SecondArgument Then
	For Each Directory in FoundDirectories
		SourceDirectory Directory, "Delete"
	Next
End If