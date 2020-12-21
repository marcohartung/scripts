'ShowDirSize.vbs
' Create a table with the dir size of all folders and subfolder
' Marco Hartung, 18.11.2020
' Version 1.0.0

'This work is licensed under the Creative Commons Attribution 4.0 International License.
'To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/ or send a letter to Creative Commons, PO Box 1866, Mountain View, CA 94042, USA.

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
	Dim fold, sEntry, i
	Dim subfolderList, subfolderDict
	
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
	
	Set subfolderDict = CreateObject("Scripting.Dictionary")
	Set subfolderList = CreateObject("System.Collections.ArrayList")
	For Each fold in fsoFolder.SubFolders
		subfolderDict.Add fold.Size, fold
	Next
	
	Dim key
	For Each key In subfolderDict.Keys
	    subfolderList.Add key
	Next
	
	subfolderList.Sort
	subfolderList.Reverse
		
	For Each key in subfolderList
		PrintFoldersizeR subfolderDict(key), fsoOutFile
	Next
	
	ActualDepth = ActualDepth - 1
End Sub
