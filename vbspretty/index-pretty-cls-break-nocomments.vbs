

Option Explicit

Dim debug : debug = (WScript.Arguments.Named("debug") = "true")
If (debug) Then WScript.Echo "Debug is enabled"
Dim VBSNEXT_TEST_INDEX : VBSNEXT_TEST_INDEX = 1
Dim vbsnextDir : vbsnextDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
Dim baseDir
With CreateObject("WScript.Shell")
	baseDir = .CurrentDirectory
End With

Public Function startsWith(str, prefix)
	startsWith = Left(str, Len(prefix)) = prefix
End Function

Public Function endsWith(str, suffix)
	endsWith = Right(str, Len(suffix)) = suffix
End Function

Public Function contains(str, char)
	contains = (InStr(1, str, char) > 0)
End Function

Public Function argsArray()
	Dim i
	ReDim arr(WScript.Arguments.Count - 1)
	For i = 0 To WScript.Arguments.Count - 1
		arr(i) = """" + WScript.Arguments(i) + """"
	Next
	argsArray = arr
End Function

Public Function argsDict()
	Dim i, param, dict
	Set dict = CreateObject("Scripting.Dictionary")
	dict.CompareMode = vbTextCompare
	ReDim arr(WScript.Arguments.Count - 1)
	For i = 1 To WScript.Arguments.Count - 1
		param = WScript.Arguments(i)
		If startsWith(param, "/") And contains(param, ":") Then
			param = Mid(param, 2)
			WScript.Echo "param to be split: " & param
			dict.Add LCase(Split(param, ":")(0)), Split(param, ":")(1)
		Else
			dict.Add i, param
		End If
	Next
	Set argsDict = dict
End Function

Class Console
	
	Public Function fmt(str, args)
		Dim res
		res = ""
		
		Dim pos
		pos = 0
		
		Dim i
		For i = 1 To Len(str)
			
			If Mid(str, i, 1) = "%" Then
				If i < Len(str) Then
					
					If Mid(str, i + 1, 1) = "%" Then
						res = res & "%"
						i = i + 1
						
					ElseIf Mid(str, i + 1, 1) = "x" Then
						res = res & CStr(args(pos))
						pos = pos + 1
						i = i + 1
					End If
				End If
				
			Else
				res = res & Mid(str, i, 1)
			End If
		Next
		
		fmt = res
	End Function
	
End Class



Dim oConsole
Set oConsole = New Console
Public Sub printf(str, args)
	
	str = Replace(str, "%s", "%x")
	str = Replace(str, "%i", "%x")
	str = Replace(str, "%f", "%x")
	str = Replace(str, "%d", "%x")
	WScript.Echo oConsole.fmt(str, args)
End Sub

Public Sub debugf(str, args)
	If (debug) Then printf str, args
End Sub

Public Sub EchoX(str, args)
	If Not IsNull(args) Then
		If IsArray(args) Then
			
			WScript.Echo oConsole.fmt(str, args)
		Else
			
			WScript.Echo oConsole.fmt(str, Array(args))
		End If
	Else
		WScript.Echo str
	End If
End Sub

Public Sub Echo(str)
	EchoX str, Null
End Sub

Public Sub EchoDX(str, args)
	If (debug) Then EchoX str, args
End Sub

Public Sub EchoD(str)
	EchoDX str, Null
End Sub

Class Collection
	
	Private dict
	Private oThis
	Private m_Name
	
	Private Sub Class_Initialize()
		Set dict = CreateObject("Scripting.Dictionary")
		Set oThis = Me
		m_Name = "Undefined"
	End Sub
	
	Public Default Property Get Obj
		Set Obj = dict
	End Property
	Public Property Set Obj(d)
		Set dict = d
	End Property
	
	Public Property Get Name
		Name = m_Name
	End Property
	Public Property Let Name(Value)
		m_Name = Value
	End Property
	
	Public Sub Add(Key, Value)
		dict.Add key, value
	End Sub
	
	Public Sub Remove(Key)
		If KeyExists(Key) Then
			dict.Remove(Key)
		Else
			RaiseErr "Key [" & Key & "] does not exists in collection."
		End If
	End Sub
	
	Public Sub RemoveAll()
		dict.RemoveAll()
	End Sub
	
	Public Property Get Count
		Count = dict.Count
	End Property
	
	Public Function GetItem(Key)
		If KeyExists(Key) Then
			GetItem = dict.Item(Key)
		Else
			
			RaiseErr "Key [" & Key & "] does not exists in collection."
		End If
	End Function
	
	Public Function GetItemAtIndex(Index)
		
		GetItemAtIndex = dict.Item(Index)
	End Function
	
	Public Function IndexOf(Key)
		IndexOf = dict.IndexOf(Key, 0)
	End Function
	
	Public Function KeyExists(Key)
		KeyExists = dict.Exists(Key)
	End Function
	
	Public Function toCSV
		toCSV = Join(toArray(), ", ")
	End Function
	
	Public Function toArray
		toArray = dict.Items
	End Function
	
	Public Function IsEmpty
		IsEmpty = (dict.Count = 0)
	End Function
	
	Private Sub RaiseErr(desc)
		Err.Clear
		Err.Raise 1000, "Collection Class Error", desc
	End Sub
	
	Private Sub Class_Terminate()
		Set dict = Nothing
		Set oThis = Nothing
	End Sub
	
End Class



Class DictUtil
	
	Function SortDictionary(objDict, intSort)
		
		Const dictKey = 1
		Const dictItem = 2
		
		Dim strDict()
		Dim objKey
		Dim strKey, strItem
		Dim X, Y, Z
		
		Z = objDict.Count
		
		If Z > 1 Then
			
			ReDim strDict(Z, 2)
			X = 0
			
			For Each objKey In objDict
				strDict(X, dictKey) = CStr(objKey)
				strDict(X, dictItem) = CStr(objDict(objKey))
				X = X + 1
			Next
			
			For X = 0 To (Z - 2)
				For Y = X To (Z - 1)
					If StrComp(strDict(X, intSort), strDict(Y, intSort), vbTextCompare) > 0 Then
						strKey = strDict(X, dictKey)
						strItem = strDict(X, dictItem)
						strDict(X, dictKey) = strDict(Y, dictKey)
						strDict(X, dictItem) = strDict(Y, dictItem)
						strDict(Y, dictKey) = strKey
						strDict(Y, dictItem) = strItem
					End If
				Next
			Next
			
			objDict.RemoveAll
			
			For X = 0 To (Z - 1)
				objDict.Add strDict(X, dictKey), strDict(X, dictItem)
			Next
			
		End If
	End Function
End Class



Class ArrayUtil
	
	Public Function toString(arr)
		If Not IsArray(arr) Then
			toString = "Supplied parameter is not an array."
			Exit Function
		End If
		
		Dim s, i
		s = "Array{" & UBound(arr) & "} [" & VbCrLf
		For i = 0 To UBound(arr)
			s = s & vbTab & "[" & i & "] => [" & arr(i) & "]"
			If i < UBound(arr) Then s = s & ", "
			s = s & VbCrLf
		Next
		s = s & "]"
		toString = s
		
	End Function
	
	Public Function contains(arr, s)
		If Not IsArray(arr) Then
			toString = "Supplied parameter is not an array."
			Exit Function
		End If
		
		Dim i, bFlag
		bFlag = False
		For i = 0 To UBound(arr)
			If arr(i) = s Then
				bFlag = True
				Exit For
			End If
		Next
		contains = bFlag
	End Function
	
End Class



Dim arrUtil
Set arrUtil = New ArrayUtil

Class PathUtil
	
	Private Property Get DOT
		DOT = "."
	End Property
	Private Property Get DOTDOT
		DOTDOT = ".."
	End Property
	
	Private oFSO
	Private m_base
	Private m_script
	Private m_temp
	
	Private Sub Class_Initialize()
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		m_script = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\") - 1)
		m_base = m_script
		m_temp = Array()
		ReDim Preserve m_temp(0)
		m_temp(0) = m_script
	End Sub
	
	Public Property Get ScriptPath
		ScriptPath = m_script
	End Property
	
	Public Property Get BasePath
		BasePath = m_base
	End Property
	
	Public Property Let BasePath(path)
		Do While endsWith(path, "\")
			path = Left(Path, Len(path) - 1)
		Loop
		m_base = Resolve(path)
		EchoDX "New Base Path: %x", m_base
	End Property
	
	Public Property Get TempBasePath
		TempBasePath = m_temp(UBound(m_temp))
	End Property
	
	Public Property Let TempBasePath(path)
		Do While endsWith(path, "\")
			path = Left(Path, Len(path) - 1)
		Loop
		If arrUtil.contains(m_temp, path) Then
			EchoDX "Temp Path %x already exists; skipped", path
		Else
			ReDim Preserve m_temp(UBound(m_temp) + 1)
			m_temp(UBound(m_temp)) = Resolve(path)
			EchoDX "New Temp Base Path: %x", m_temp(UBound(m_temp))
		End If
	End Property
	
	Function Resolve(path)
		Dim pathBase, lPath, final
		EchoDX "path: %x", path
		If path = DOT Or path = DOTDOT Then
			path = path & "\"
		End If
		EchoDX "path: %x", path
		
		If oFSO.FolderExists(path) Then
			EchoD "FolderExists"
			Resolve = oFSO.GetFolder(path).path
			Exit Function
		End If
		
		If oFSO.FileExists(path) Then
			EchoD "FileExists"
			Resolve = oFSO.GetFile(path).path
			Exit Function
		End If
		
		pathBase = oFSO.BuildPath(m_base, path)
		EchoDX "Adding base %x to path %x. New Path: %x", Array(m_base, path, pathBase)
		
		If endsWith(pathBase, "\") Then
			If IsObject(oFSO.GetFolder(pathBase)) Then
				EchoD "EndsWith '\' -> FolderExists"
				Resolve = oFSO.GetFolder(pathBase).Path
				Exit Function
			End If
		Else
			
			If oFSO.FolderExists(pathBase) Then
				EchoD "FolderExists"
				Resolve = oFSO.GetFolder(pathBase).path
				Exit Function
			End If
			
			If oFSO.FileExists(pathBase) Then
				EchoD "FileExists"
				Resolve = oFSO.GetFile(pathBase).path
				Exit Function
			End If
			
			Dim i
			i = UBound(m_temp)
			Do
				lPath = oFSO.BuildPath(m_temp(i), path)
				EchoDX "Adding Temp Base path (%x) %x to path %x. New Path: %x", Array(i, m_temp(i), path, lPath)
				If oFSO.FileExists(lPath) Then
					final = oFSO.GetFile(lPath).path
					EchoDX "File Resolved with Temp Base %x", final
					Resolve = final
					Exit Function
				End If
				If oFSO.FolderExists(lPath) Then
					final = oFSO.GetFolder(lPath)
					EchoDX "Folder Resolved with Temp Base %x", final
					Resolve = final
					Exit Function
				End If
				i = i - 1
			Loop While i >= 0
			
			lPath = oFSO.BuildPath(m_script, path)
			EchoDX "Adding script path %x to path %x. New Path: %x", Array(m_script, path, lPath)
			If oFSO.FileExists(lPath) Then
				final = oFSO.GetFile(lPath).path
				EchoDX "File Resolved with Temp Base %x", final
				Resolve = final
				Exit Function
			End If
			If oFSO.FolderExists(lPath) Then
				final = oFSO.GetFolder(lPath)
				EchoDX "Folder Resolved with Temp Base %x", final
				Resolve = final
				Exit Function
			End If
		End If
		
		EchoD "Unable to Resolve"
		Resolve = path
	End Function
	
	Private Sub Class_Terminate()
		Set oFSO = Nothing
	End Sub
	
End Class



Dim putil
Set putil = New PathUtil
putil.BasePath = baseDir
EchoX "Project location: %x", putil.BasePath

Class FSO
	Private dir
	Private objFSO
	
	Private Sub Class_Initialize
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		dir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
	End Sub
	
	Public Sub setDir(s)
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
	
	Public Function CreateFolder(fol)
		CreateFolder = False
		If FolderExists(fol) Then
			CreateFolder = True
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
		Set objFile = objFSO.GetFile(file)
		On Error GoTo 0
		If IsObject(objFile) Then
			GetFilePath = objFile.Path
		Else
			EchoDX "File %x not found; searching in directory %x", Array(file, dir)
			On Error Resume Next
			Set objFile = objFile.GetFile(objFSO.BuildPath(dir, file))
			On Error GoTo 0
			If IsObject(objFile) Then
				GetFilePath = objFile.Path
			Else
				GetFilePath = "File [" & file & "] Not found"
			End If
		End If
	End Function
	
	Public Function GetFileName(ByVal file)
		GetFileName = objFSO.GetFile(file).Name
	End Function
	
	Public Function GetFileExtn(file)
		GetFileExtn = ""
		On Error Resume Next
		GetFileExtn = LCase(objFSO.GetExtensionName(file))
		On Error GoTo 0
	End Function
	
	Public Function GetBaseName(ByVal file)
		GetBaseName = Replace(GetFileName(file), "." & GetFileExtn(file), "")
	End Function
	
	Public Function ReadFile(file)
		file = putil.Resolve(file)
		EchoDX "---> File resolved to: %x", Array(file)
		If Not FileExists(file) Then
			Wscript.Echo "---> File " & file & " does not exists."
			ReadFile = ""
			Exit Function
		End If
		Dim objFile : Set objFile = objFSO.OpenTextFile(file)
		ReadFile = objFile.ReadAll()
		objFile.Close
	End Function
	
	Public Function FileExists(file)
		FileExists = objFSO.FileExists(file)
	End Function
	
	Public Sub DeleteFile(file)
		On Error Resume Next
		objFSO.DeleteFile(file)
		On Error GoTo 0
	End Sub
	
End Class



Dim cFS
Set cFS = New FSO

cFS.setDir(baseDir)

Public Function Log(msg)
	cFS.WriteFile "build.log", msg, False
End Function

Log "VBSNext Directory: " & vbsnextDir

Class ClassA
	Public Default Sub CallMe
		WScript.Echo "Class-extending resolved successfully."
	End Sub
End Class



Class ClassB
	
	Private m_CLASSA
	
	Private Sub Class_Initialize
		Set m_CLASSA = New CLASSA
	End Sub
	
	Public Default Sub CallMe
		Call m_CLASSA.CallMe
	End Sub
End Class



Dim ccb
Set ccb = New ClassB
ccb.CallMe

Public Sub Include(file)
	
End Sub
Public Sub Import(file)
	
End Sub



Class VbsJson



	Private Whitespace, NumberRegex, StringChunk
	Private b, f, r, n, t
	
	Private Sub Class_Initialize
		Whitespace = " " & vbTab & vbCr & vbLf
		b = ChrW(8)
		f = vbFormFeed
		r = vbCr
		n = vbLf
		t = vbTab
		
		Set NumberRegex = New RegExp
		NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
		NumberRegex.Global = False
		NumberRegex.MultiLine = True
		NumberRegex.IgnoreCase = True
		
		Set StringChunk = New RegExp
		StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
		StringChunk.Global = False
		StringChunk.MultiLine = True
		StringChunk.IgnoreCase = True
	End Sub




















	Public Function Encode(ByRef obj)
		Dim buf, i, c, g
		Set buf = CreateObject("Scripting.Dictionary")
		Select Case VarType(obj)
			Case vbNull
				buf.Add buf.Count, "null"
			Case vbBoolean
				If obj Then
					buf.Add buf.Count, "true"
				Else
					buf.Add buf.Count, "false"
				End If
			Case vbInteger, vbLong, vbSingle, vbDouble
				buf.Add buf.Count, obj
			Case vbString
				buf.Add buf.Count, """"
				For i = 1 To Len(obj)
					c = Mid(obj, i, 1)
					Select Case c
						Case """" buf.Add buf.Count, "\"""
						Case "\" buf.Add buf.Count, "\\"
						Case "/" buf.Add buf.Count, "/"
						Case b buf.Add buf.Count, "\b"
						Case f buf.Add buf.Count, "\f"
						Case r buf.Add buf.Count, "\r"
						Case n buf.Add buf.Count, "\n"
						Case t buf.Add buf.Count, "\t"
						Case Else
							If AscW(c) >= 0 And AscW(c) <= 31 Then
								c = Right("0" & Hex(AscW(c)), 2)
								buf.Add buf.Count, "\u00" & c
							Else
								buf.Add buf.Count, c
							End If
					End Select
				Next
				buf.Add buf.Count, """"
			Case vbArray + vbVariant
				g = True
				buf.Add buf.Count, "["
				For Each i In obj
					If g Then g = False Else buf.Add buf.Count, ","
					buf.Add buf.Count, Encode(i)
				Next
				buf.Add buf.Count, "]"
			Case vbObject
				If TypeName(obj) = "Dictionary" Then
					g = True
					buf.Add buf.Count, "{"
					For Each i In obj
						If g Then g = False Else buf.Add buf.Count, ","
						buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
					Next
					buf.Add buf.Count, "}"
				Else
					Err.Raise 8732, , "None dictionary object"
				End If
			Case Else
				buf.Add buf.Count, """" & CStr(obj) & """"
		End Select
		Encode = Join(buf.Items, "")
	End Function
	



















	Public Function Decode(ByRef str)
		Dim idx
		idx = SkipWhitespace(str, 1)
		
		If Mid(str, idx, 1) = "{" Then
			Set Decode = ScanOnce(str, 1)
		Else
			Decode = ScanOnce(str, 1)
		End If
	End Function

	Private Function ScanOnce(ByRef str, ByRef idx)
		Dim c, ms
		
		idx = SkipWhitespace(str, idx)
		c = Mid(str, idx, 1)
		
		If c = "{" Then
			idx = idx + 1
			Set ScanOnce = ParseObject(str, idx)
			Exit Function
		ElseIf c = "[" Then
			idx = idx + 1
			ScanOnce = ParseArray(str, idx)
			Exit Function
		ElseIf c = """" Then
			idx = idx + 1
			ScanOnce = ParseString(str, idx)
			Exit Function
		ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ScanOnce = Null
			Exit Function
		ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ScanOnce = True
			Exit Function
		ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
			idx = idx + 5
			ScanOnce = False
			Exit Function
		End If

		Set ms = NumberRegex.Execute(Mid(str, idx))
		If ms.Count = 1 Then
			idx = idx + ms(0).Length
			ScanOnce = CDbl(ms(0))
			Exit Function
		End If

		Err.Raise 8732, , "No JSON object could be ScanOnced"
	End Function
	
	Private Function ParseObject(ByRef str, ByRef idx)
		Dim c, key, value
		Set ParseObject = CreateObject("Scripting.Dictionary")
		idx = SkipWhitespace(str, idx)
		c = Mid(str, idx, 1)

		If c = "}" Then
			Exit Function
		ElseIf c <> """" Then
			Err.Raise 8732, , "Expecting property name"
		End If
		
		idx = idx + 1

		Do
			key = ParseString(str, idx)
			
			idx = SkipWhitespace(str, idx)
			If Mid(str, idx, 1) <> ":" Then
				Err.Raise 8732, , "Expecting : delimiter"
			End If
			
			idx = SkipWhitespace(str, idx + 1)
			If Mid(str, idx, 1) = "{" Then
				Set value = ScanOnce(str, idx)
			Else
				value = ScanOnce(str, idx)
			End If
			ParseObject.Add key, value
			
			idx = SkipWhitespace(str, idx)
			c = Mid(str, idx, 1)
			If c = "}" Then
				Exit Do
			ElseIf c <> "," Then
				Err.Raise 8732, , "Expecting , delimiter"
			End If
			
			idx = SkipWhitespace(str, idx + 1)
			c = Mid(str, idx, 1)
			If c <> """" Then
				Err.Raise 8732, , "Expecting property name"
			End If
			
			idx = idx + 1
		Loop
		
		idx = idx + 1
	End Function

	Private Function ParseArray(ByRef str, ByRef idx)
		Dim c, values, value
		Set values = CreateObject("Scripting.Dictionary")
		idx = SkipWhitespace(str, idx)
		c = Mid(str, idx, 1)
		
		If c = "]" Then
			ParseArray = values.Items
			Exit Function
		End If
		
		Do
			idx = SkipWhitespace(str, idx)
			If Mid(str, idx, 1) = "{" Then
				Set value = ScanOnce(str, idx)
			Else
				value = ScanOnce(str, idx)
			End If
			values.Add values.Count, value
			
			idx = SkipWhitespace(str, idx)
			c = Mid(str, idx, 1)
			If c = "]" Then
				Exit Do
			ElseIf c <> "," Then
				Err.Raise 8732, , "Expecting , delimiter"
			End If
			
			idx = idx + 1
		Loop
		
		idx = idx + 1
		ParseArray = values.Items
	End Function

	Private Function ParseString(ByRef str, ByRef idx)
		Dim chunks, content, terminator, ms, esc, char
		Set chunks = CreateObject("Scripting.Dictionary")
		
		Do
			Set ms = StringChunk.Execute(Mid(str, idx))
			If ms.Count = 0 Then
				Err.Raise 8732, , "Unterminated string starting"
			End If

			content = ms(0).Submatches(0)
			terminator = ms(0).Submatches(1)
			If Len(content) > 0 Then
				chunks.Add chunks.Count, content
			End If

			idx = idx + ms(0).Length

			If terminator = """" Then
				Exit Do
			ElseIf terminator <> "\" Then
				Err.Raise 8732, , "Invalid control character"
			End If

			esc = Mid(str, idx, 1)
			
			If esc <> "u" Then
				Select Case esc
					Case """" char = """"
					Case "\" char = "\"
					Case "/" char = "/"
					Case "b" char = b
					Case "f" char = f
					Case "n" char = n
					Case "r" char = r
					Case "t" char = t
					Case Else Err.Raise 8732, , "Invalid escape"
				End Select
				idx = idx + 1
			Else
				char = ChrW("&H" & Mid(str, idx + 1, 4))
				idx = idx + 5
			End If
			
			chunks.Add chunks.Count, char
		Loop
		
		ParseString = Join(chunks.Items, "")
	End Function
	
	Private Function SkipWhitespace(ByRef str, ByVal idx)
		Do While idx <= Len(str) And _
				InStr(Whitespace, Mid(str, idx, 1)) > 0
			idx = idx + 1
		Loop
		SkipWhitespace = idx
	End Function
	
End Class




Include("classes\VbsJson")
Dim json, str, o, i, k

Set json = New VbsJson
str = "{""keys"":[1,""a""]}"
Set o = json.Decode(str)
For Each k In o("keys")
	WScript.Echo k
Next





str = cfs.ReadFile(".\data\data.json")
Set o = json.Decode(str)
WScript.Echo o("Image")("Width")
WScript.Echo o("Image")("Height")
WScript.Echo o("Image")("Title")
WScript.Echo o("Image")("Thumbnail")("Url")
For Each i In o("Image")("IDs")
	WScript.Echo i
Next



Class Person
	Private m_Age
	Private m_Name
	
	Public Default Function Init(Name, Age)
		m_Name = Name
		m_Age = Age

		Set Init = Me
	End Function

	Public Property Get Name
		Name = m_Name
	End Property
	Public Property Let Name(v)
		m_Name = v
	End Property

	Public Property Get Age
		Age = m_Age
	End Property
	Public Property Let Age(v)
		m_Age = v
	End Property
	
	Public Property Get toString
		toString = m_Name & " (" & m_Age & ")"
	End Property
End Class



Include ".\classes\Person"

Dim TheDude : Set TheDude = (New Person)("John", 40)
WScript.Echo TheDude.toString





Include("lib\person-test")
Include("lib\json-test")