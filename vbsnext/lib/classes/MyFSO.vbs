' ==============================================================================================
' Implementation of several use cases of FileSystemObject into this class
' Author: Praveen Nandagiri (pravynandas@gmail.com)
' ==============================================================================================

Class MyFSO
	Private dir
	Private objFSO
	Private pUtil
	
	Private Sub Class_Initialize
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		dir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
		set pUtil = new PathUtil
		pUtil.BasePath = dir
	End Sub

	' Update the current directory of the instance if needed
	public Sub setDir(s)
		dir = s
	End Sub

	Public Function getDir
		getDir = dir
	End Function
	
	Public Function GetFSO
		Set GetFSO = objFSO
	End Function

  Public Function FolderExists(fol)
    FolderExists = objFSO.FolderExists(fol)
  End Function
    ' ===================== Sub Routines =====================


	Public Function CreateFolder(fol)
    CreateFolder = false
		If FolderExists(fol) Then
      CreateFolder = true
    Else
			objFSO.CreateFolder(fol)
      CreateFolder = FolderExists(fol)
		End If
	End Function
	
	Public Sub WriteFile(strFileName, strMessage, overwrite)
		Const ForReading = 1
		Const ForWriting = 2
		Const ForAppending = 8
		Dim mode
    	Dim oFile
		
    	mode = ForWriting
		If Not overwrite Then
			mode = ForAppending
		End If
		
		If objFSO.FileExists(strFileName) Then
			Set oFile = objFSO.OpenTextFile(strFileName, mode)
		Else
			Set oFile = objFSO.CreateTextFile(strFileName)
		End If
		oFile.WriteLine strMessage
		
		oFile.Close
		
		Set oFile = Nothing
	End Sub 

	' ===================== Function Routines =====================

	Public Function GetFileDir(ByVal file)
    	EchoDX "GetFileDir( %x )", Array(file)
		Dim objFile
		Set objFile = objFSO.GetFile(file)
		GetFileDir = objFSO.GetParentFolderName(objFile) 
	End Function
	
	Public Function GetFilePath(ByVal file)
    EchoDX "GetFilePath( %x )", Array(file)
    Dim objFile
    On Error Resume Next
    set objFile = objFSO.GetFile(file)
    On Error Goto 0
    If IsObject(objFile) Then
		  GetFilePath = objFile.Path 
    Else
      EchoDX "File %x not found; searching in directory %x", Array(file,dir)
      On Error Resume Next
      set objFile = objFile.GetFile(objFSO.BuildPath(dir, file))
      On Error Goto 0
      If IsObject(objFile) Then
		    GetFilePath = objFile.Path 
      Else
        GetFilePath = "File [" & file & "] Not found"
      End If
    End If
	End Function

  ''' <summary>Returns a specified number of characters from a string.</summary>
  ''' <param name="file">File Name</param>
	Public Function GetFileName(ByVal file)
		GetFileName = objFSO.GetFile(file).Name
	End Function

	Public Function GetFileExtn(file)
		GetFileExtn = ""
		on Error Resume Next
		GetFileExtn = LCASE(objFSO.GetExtensionName(file))
		On Error goto 0
	End Function

  Public Function GetBaseName(ByVal file)
    GetBaseName = Replace(GetFileName(file), "." & GetFileExtn(file), "")
  End Function

	'TODO: This resolves files only locally; If a package imported in another project, paths won't resolve. 
	' Example, .\data.json
	Public Function ReadFile(file)
		file = pUtil.Resolve(file)
		If Not FileExists(file) Then 
			Wscript.Echo "File " & file & " does not exists."
			ReadFile = ""
			Exit Function
		End If
		Dim objFile: Set objFile = objFSO.OpenTextFile(file)
		ReadFile = objFile.ReadAll()
		objFile.Close
	End Function

	Public Function FileExists(file)
		FileExists = objFSO.FileExists(file)
	End Function

	Public Sub DeleteFile(file)
		on Error resume next
		objFSO.DeleteFile(file)
		On Error Goto 0
	End Sub


End Class