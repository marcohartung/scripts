'ShowDirSize.vbs
' Create a table with the dir size of all folders and subfolder
' Marco Hartung, 18.11.2020

' "WshShell.Run strScript, 0" 0 = hide window, 1 = show window (useful for debugging)
Option Explicit

Dim fso, sFolder, fsoOutFile, strOutFile, fsoFolder
Dim MaxDepth, ActualDepth

ActualDepth = 0
MaxDepth = 2

Set fso = CreateObject( "Scripting.FileSystemObject" )

sFolder = Wscript.Arguments.Item(0)
 'sFolder = "C:\Temp"
If sFolder = "" Then
  Wscript.Echo "No Folder parameter was passed"
  Wscript.Quit
End If

If Not( fso.FolderExists(sFolder) ) then
Wscript.Echo "Folder did not exist: " & sFolder
  Wscript.Quit
End If

strOutFile = fso.BuildPath("C:\Temp", "FolderSizeInfo.txt" )
Set fsoOutFile = fso.CreateTextFile(strOutFile) 
Set fsoFolder = fso.GetFolder(sFolder)

PrintFoldersizeR fsoFolder, fsoOutFile

wscript.echo "Finished!"

' -------------------------- functions  --------------------------

Sub PrintFoldersizeR( fsoFolder,  fsoOutFile )
	Dim fold, sEntry
	
	If( ActualDepth > MaxDepth) Then
		Exit Sub
	End If
	ActualDepth = ActualDepth + 1
	

	sEntry = _ 
		fsoFolder.Path & _
		vbTab & _
		Round( fsoFolder.Size / (1024*1024), 2) & _
		"MB"
	
	fsoOutFile.WriteLine sEntry 
	
	For Each fold in fsoFolder.SubFolders
		PrintFoldersizeR fold, fsoOutFile
	Next
	
	ActualDepth = ActualDepth - 1

End Sub
