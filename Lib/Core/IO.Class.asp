<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统文件操作类]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

Class Cls_IO
	
	'// 定义私有命名对象
	Private PrFSO, PrCAS
	
	'// 定义FSO名称，以及操作文件的编码
	Private PrName, PrCharset
	
	'// 初始化类
	Private Sub Class_Initialize()
		'// 初始化FSO对象名称
		PrName = "Scripting.FileSystemObject"
		
		'// 初始化编码格式，继承自主类
		PrCharset = System.Charset
		
		'// 初始化系统默认的错误信息
		System.Error.E(52) = "写入文件错误！"
		System.Error.E(53) = "创建文件夹错误！"
		System.Error.E(54) = "读取文件列表失败！"
		System.Error.E(55) = "设置属性失败，文件不存在！"
		System.Error.E(56) = "设置属性失败！"
		System.Error.E(57) = "获取属性失败，文件不存在！"
		System.Error.E(58) = "复制失败，源文件不存在！"
		System.Error.E(59) = "移动失败，源文件不存在！"
		System.Error.E(60) = "删除失败，文件不存在！"
		System.Error.E(61) = "重命名失败，源文件不存在！"
		System.Error.E(62) = "重命名失败，已存在同名文件！"
		System.Error.E(63) = "文件或文件夹操作错误！"
	End Sub
	
	'// 释放类
	Private Sub Class_Terminate()
		If IsObject(PrFSO) Then Set PrFSO = Nothing End If
		If IsObject(PrCAS) Then Set PrCAS = Nothing End If
	End Sub
	
	'// 声明文件操作对象模块单元
	Public Property Get FSO()
		If Not IsObject(PrFSO) Then Set PrFSO = Server.CreateObject(PrName) End If
		Set FSO = PrFSO
	End Property
	Public Property Get CAS()
		If Not IsObject(PrCAS) Then Set PrCAS = New Cls_IO_CAS End If
		Set CAS = PrCAS
	End Property
	
	'// 设置服务器FSO组件的名称
	Public Property Let Name(ByVal blParam)
		PrName = blParam
	End Property
	
	'// 设置操作文件的编码
	Public Property Let Charset(ByVal blParam)
		PrCharset = blParam
	End Property
	
	'/**
	' * @功能说明: 动态包含文件
	' * @参数说明: - blFilePath [string]: 目标文件路径
	' */
	Public Sub Import(ByVal blFilePath)
		ExecuteGlobal ReadInclude(blFilePath, 0)
	End Sub
	
	'// 获取文件内容
	Private Function ReadInclude(ByVal blFilePath, ByVal blHtml)
		Dim blContentStartPosition, blCodeStartPosition
		Dim blContent, blTempContent, blCode, blTempCode, blHtmlCode
		blContent = ReadIncludes(blFilePath)
		blCode = "": blContentStartPosition = 1: blCodeStartPosition = InStr(blContent, "<"&"%") + 2
		blHtmlCode = System.Text.IIF(blHtml = 1, "blACLHtml = blACLHtml & ","Response.Write ")
		While blCodeStartPosition > blContentStartPosition + 1
			blTempContent = Mid(blContent, blContentStartPosition, blCodeStartPosition - blContentStartPosition - 2)
			blContentStartPosition = InStr(blCodeStartPosition, blContent, "%"&">") + 2
			If Not System.Text.IsEmptyAndNull(blTempContent) Then
				blTempContent = Replace(blTempContent, """", """""")
				blTempContent = Replace(blTempContent, vbCrLf&vbCrLf, vbCrLf)
				blTempContent = Replace(blTempContent, vbCrLf, """&vbCrLf&""")
				blCode = blCode & blHtmlCode & """" & blTempContent & """" & vbCrLf
			End If
			blTempContent = Mid(blContent, blCodeStartPosition, blContentStartPosition - blCodeStartPosition - 2)
			blTempCode = System.Text.ReplaceX(blTempContent, "^\s*=\s*", blHtmlCode) & vbCrLf
			If blHtml = 1 Then
				blTempCode = System.Text.ReplaceXMultiline(blTempCode, "^(\s*)Response\.Write", "$1" & blHtmlCode) & vbCrLf
				blTempCode = System.Text.ReplaceXMultiline(blTempCode, "^(\s*)System\.(WB|W|WE|WR)", "$1" & blHtmlCode) & vbCrLf
			End If
			blCode = blCode & Replace(blTempCode, vbCrLf&vbCrLf, vbCrLf)
			blCodeStartPosition = InStr(blContentStartPosition, blContent, "<"&"%") + 2
		Wend
		blTempContent = Mid(blContent,blContentStartPosition)
		If Not System.Text.IsEmptyAndNull(blTempContent) Then
			blTempContent = Replace(blTempContent, """", """""")
			blTempContent = Replace(blTempContent, vbCrLf&vbCrLf, vbCrLf)
			blTempContent = Replace(blTempContent, vbCrlf,"""&vbCrLf&""")
			blCode = blCode & blHtmlCode & """" & blTempContent & """" & vbCrLf
		End If
		If blHtml = 1 Then blCode = "blACLHtml = """" " & vbCrLf & blCode
		ReadInclude = Replace(blCode, vbCrLf&vbCrLf, vbCrLf)
	End Function
	
	'// 递归获取包含文件的内容
	Private Function ReadIncludes(ByVal blFilePath)
		Dim blContent: blContent = Me.Read(blFilePath)
		If Not System.Text.IsEmptyAndNull(blContent) Then
			blContent = System.Text.ReplaceX(blContent, "<"&"% *?@.*?%"&">", "")
			blContent = System.Text.ReplaceX(blContent, "(<"&"%[^>]+?)(option +?explicit)([^>]*?%"&">)", "$1'$2$3")
			Dim blRule: blRule = "<!-- *?#include +?(file|virtual) *?= *?""??([^"":?*\f\n\r\t\v]+?)""?? *?-->"
			'// 判断文件中是否有包含其他文件
			If System.Text.Test(blContent, blRule) Then
				Dim blIncludeFile, blIncludeFileContent
				Dim blMatches: Set blMatches = System.Text.MatchX(blContent, blRule)
				Dim blMatch: For Each blMatch In blMatches
					If LCase(blMatch.SubMatches(0)) = "virtual" Then blIncludeFile = blMatch.SubMatches(1) _
					Else blIncludeFile = Mid(blFilePath, 1, InstrRev(blFilePath, System.Text.IIF(Instr(blFilePath, ":") > 0, "\", "/"))) & blMatch.SubMatches(1)
					'// 递归获取包含文件的内容
					blIncludeFileContent = ReadIncludes(blIncludeFile)
					blContent = Replace(blContent, blMatch, blIncludeFileContent)
				Next
				Set blMatches = Nothing
			End If
		End If
		ReadIncludes = blContent
	End Function
		
	'/**
	' * @功能说明: 读取指定文件对象
	' * @参数说明: - blFile [string]: 对象路径
	' * @返回值:   - [string] 字符串
	' */
	Public Function Read(ByVal blFile)
		If ExistsFile(blFile) Then
			blFile = FormatFilePath(blFile)
			Dim blStream: Set blStream = Server.CreateObject("ADODB.Stream")
			With blStream
				.Type = 2: .Mode = 3
				.Charset = PrCharset: .Open
				.LoadFromFile blFile: Read = .ReadText
				.Close
			End With
			Set blStream = Nothing
		Else Read = "" End If
	End Function

	'/**
	' * @功能说明: 覆盖当前打开的文本内容，文件及文件夹不存在则创建
	' * @参数说明: - blFile [string]: 对象路径
	' * 		   - blContent [string]: 保存内容
	' * @返回值:   - [bool]: 布尔值
	' */
	Public Function Save(ByVal blFile, ByVal blContent)
		If Not System.Text.IsEmptyAndNull(blFile) Then
			blFile = FormatFilePath(blFile)
			Dim blFolder: blFolder = Directory(blFile, "\")

			'// 如果文件夹不存在，则创建新文件夹
			If Not ExistsFolder(blFolder) Then CreateFolder(blFolder)
			'// 只有在文件夹存在时，对文件进行保存
			If ExistsFolder(blFolder) Then
				On Error Resume Next
				Dim blStream: Set blStream = Server.CreateObject("ADODB.Stream")
				With blStream
					.Open
					.Charset = PrCharset
					.Position = blStream.Size
					.WriteText = blContent
					.SaveToFile blFile, 2
					.Close
				End With
				If Err Then
					System.Error.Message = "（"& blFile &"）"
					System.Error.Raise 52
				End If
				Err.Clear
				Save = True: Set blStream = Nothing
			Else Save = False End If
		Else Save = False End If
	End Function
	
	'/**
	' * @功能说明: 删除文件(同时支持绝对和相对两种路径模式)
	' * @参数说明: - blFile [string]: 删除文件对象
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function Delete(ByVal blFile)		
		If Not ExistsFile(blFile) Then Delete = False _
		Else FSO.DeleteFile FormatFilePath(blFile): Delete = True
	End Function
	
	'/**    
	' * @功能说明： 遍历目录下的所有目录和文件（不包括子目录）
	' * @参数说明： - [string] blPath : 初始路径    
	' * 			- [bool] blShowFile : 是否遍历文件
	' * @返回值：   - [array] : 二维数组
	' */	
	Public Function Dir(ByVal blPath, ByVal blShowFile)
		If ExistsFolder(blPath) Then
			Dim blArray(), I: I = 0
			Dim blFolder, blSubFolder, blItem
			Set blFolder = FSO.GetFolder(FormatFilePath(blPath))
			Set blSubFolder = blFolder.SubFolders
			ReDim Preserve blArray(4, blSubFolder.Count - 1)
			For Each blItem In blSubFolder
				blArray(0, I) = blItem.Name & "/"
				blArray(1, I) = blItem.Size
				blArray(2, I) = blItem.DateLastModified
				blArray(3, I) = blItem.Attributes
				blArray(4, I) = blItem.Type
				I = I + 1
			Next
			'// 判断是否显示文件
			If System.Text.ToBoolean(blShowFile) Then
				Set blSubFolder = blFolder.Files
				ReDim Preserve blArray(4, blSubFolder.Count + I - 1)
				For Each blItem In blSubFolder
					blArray(0, I) = blItem.Name
					blArray(1, I) = blItem.Size
					blArray(2, I) = blItem.DateLastModified
					blArray(3, I) = blItem.Attributes
					blArray(4, I) = blItem.Type
					I = I + 1
				Next
			End If
			Set blSubFolder = Nothing
			Set blFolder = Nothing
			Dir = blArray
		Else ReDim blArray2(-1, -1): Dir = blArray2 End If
	End Function
	
	'/**    
	' * @功能说明： 遍历目录下的所有目录和文件（包括子目录）
	' * @参数说明： - [string] sPath : 初始路径    
	' *  			- [bool] bAll : 是否遍历子目录
	' * 			- [bool] bFile : 是否遍历文件
	' * @返回值：   - [array] : 数组
	' */    
	Public Function Traversal(ByVal sPath, ByVal bAll, ByVal bFile)
		If ExistsFolder(sPath) Then
			Dim oKey, pKey, nKey, mKey
			Dim mItem, mName, nItem, nPath
			Dim oDic, oArray
			Set oDic = Server.CreateObject("Scripting.Dictionary")
			For Each nItem In FSO.GetFolder(FormatFilePath(sPath)).SubFolders
				nPath = sPath & nItem.Name & "/"
				oKey = System.Security.MD5(nPath, 16)
				If Not oDic.Exists(oKey) Then oDic.Add oKey, nPath
				If System.Text.ToBoolean(bFile) Then
					For Each mItem In nItem.Files
						mName = nPath & mItem.Name
						nKey = System.Security.MD5(mName, 16)
						If Not oDic.Exists(nKey) Then oDic.Add nKey, mName
					Next
				End If
				
				If System.Text.ToBoolean(bAll) Then
					If System.Text.ToBoolean(bFile) Then oArray = Traversal(nPath, True, True) _
					Else oArray = Traversal(nPath, True, False)
					
					Dim I: For I = 0 To UBound(oArray)
						pKey = System.Security.MD5(oArray(I), 16)
						If Not oDic.Exists(pKey) Then oDic.Add pKey, oArray(I)
						If System.Text.ToBoolean(bFile) Then
							For Each mItem In nItem.Files
								mName = nPath & mItem.Name
								mKey = System.Security.MD5(mName, 16)
								If Not oDic.Exists(mKey) Then oDic.Add mKey, mName
							Next
						End If
					Next
				End If
			Next
			Traversal = oDic.Items
			Set oDic = Nothing
		Else ReDim blArray(-1): Traversal = blArray End If
	End Function

	'/**
	' * @功能说明: 新建文件夹（同时支持绝对和相对两种路径模式）
	' * @参数说明: - blFolder [string]: 新建文件夹对象
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function CreateFolder(ByVal blFolder)
		If Not System.Text.IsEmptyAndNull(blFolder) Then
			Dim I, ArrayPaths
			Dim blPaths: blPaths = Split(FormatFilePath(blFolder), "\")
			For I = 0 To UBound(blPaths)
				If I = 0 Then ArrayPaths = blPaths(I) Else ArrayPaths = ArrayPaths & "\" & blPaths(I)
				If I > 0 Then
					'// 当前文件夹下，如果有文件与文件夹同名时，将无法创建文件夹。
					If ExistsFile(ArrayPaths) Then
						System.Error.Message = "("& ArrayPaths &")"
						System.Error.Raise 53
						CreateFolder = False: Exit Function
					Else
						If Not ExistsFolder(ArrayPaths) Then FSO.CreateFolder ArrayPaths End If
					End If
				End If
			Next
			CreateFolder = True
		Else CreateFolder = False End If
	End Function

	'/**
	' * @功能说明: 删除文件夹(同时支持绝对和相对两种路径模式)
	' * @参数说明: - blFolder [string]: 删除文件夹对象
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function DeleteFolder(ByVal blFolder)		
		If Not System.Text.IsEmptyAndNull(blFolder) Then
			If ExistsFolder(blFolder) Then FSO.DeleteFolder FormatFilePath(blFolder): DeleteFolder = True _
			Else DeleteFolder = False
		Else DeleteFolder = False End If
	End Function

	'/**
	' * @功能说明: 检查文件夹是否存在
	' * @参数说明: - blFolder [string]: 对象路径
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function ExistsFolder(ByVal blFolder)
		If System.Text.IsEmptyAndNull(blFolder) Then ExistsFolder = False _
		Else ExistsFolder = FSO.FolderExists(FormatFilePath(blFolder))
	End Function

	'/**
	' * @功能说明: 检查文件是否存在
	' * @参数说明: - blFile [string]: 检测对象的路径
	' * @返回值:   - [bool] 布尔值
	' */
	Public Function ExistsFile(ByVal blFile)
		If System.Text.IsEmptyAndNull(blFile) Then ExistsFile = False _
		Else ExistsFile = FSO.FileExists(FormatFilePath(blFile))
	End Function

	'/**
	' * @功能说明: 将目标文件（相对路径）转换为绝对路径
	' * @参数说明: - blFile [string]: 目标文件路径
	' * @返回值:   - [string] 字符串
	' * @函数说明: 函数先转换所出现的错误字符。基于：Errorchar变量
	' *            判断当前输入参数是绝对路径或相对路径
	' *            如果是绝对路径，替换所有的相对路径所使用的字符
	' */
	Public Function FormatFilePath(ByVal blFile)
		Dim I, blParam3: blParam3 = Empty
		Dim blParam2: blParam2 = Trim(blFile)
		Dim blIllegal: blIllegal = "',"",*,?,&,|,<,>,;"
		Dim blArrayIllegal: blArrayIllegal = Split(blIllegal, ",")		
		For I = 0 To UBound(blArrayIllegal)
			If InStr(blParam2, blArrayIllegal(I)) > 0 Then
				blParam2 = Replace(blParam2, blArrayIllegal(I), "")
			End If
		Next
		
		blParam2 = System.Text.ReplaceX(blParam2, "(\/|\\)+", "/")
		
		'// 判断目标路径是否为绝对路径
		If Mid(blParam2, 2, 1) <> ":" Then blParam2 = Server.MapPath(blParam2) _
		Else blParam2 = Replace(blParam2, "/", "\")
		
		FormatFilePath = System.Text.IIF(Right(blParam2, 1) = "\", Left(blParam2, Len(blParam2) - 1), blParam2)
	End Function

	'// 取文件夹绝对路径
	Private Function absPath(ByVal p)
		If System.Text.IsEmptyAndNull(p) Then absPath = "" : Exit Function
		If Mid(p, 2, 1) <> ":" Then
			If isWildcards(p) Then
				p = Replace(p, "*", "[.$.[a.c.l.s.t.a.r].#.]")
				p = Replace(p, "?", "[.$.[a.c.l.q.u.e.s].#.]")
				p = Server.MapPath(p)
				p = Replace(p, "[.$.[a.c.l.q.u.e.s].#.]", "?")
				p = Replace(p, "[.$.[a.c.l.s.t.a.r].#.]", "*")
			Else
				p = Server.MapPath(p)
			End If
		End If
		If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
		absPath = p
	End Function
	
	'// 路径是否包含通配符
	Private Function isWildcards(ByVal path)
		isWildcards = False
		If InStr(path, "*") > 0 Or InStr(path, "?") > 0 Then isWildcards = True
	End Function
	
	'/**
	' * @功能说明: 获取指定文件的目录
	' * @参数说明: - blFile [string]: 目标文件路径
	' * 		   - blParam2 [string]: 查询关键字，默认为"/"
	' * @返回值:   - [string]: 字符串
	' */
	Public Function Directory(ByVal blFile, ByVal blParam2)
		Dim blSplits: blSplits = System.Text.IIF(System.Text.IsEmptyAndNull(blParam2), "/", blParam2)
		Directory = System.Text.IIF(InStrRev(blFile, blSplits) < 1, "/", Mid(blFile, 1, InStrRev(blFile, blSplits)))
	End Function

	'/**
	' * @功能说明: 获取文件的后缀名
	' * @参数说明: - blFile [string]: 目标文件
	' * @返回值:   - [string] 字符串
	' */
	Public Function FileExts(ByVal blFile)
		FileExts = "Unknow"
		FileExts = LCase(Split(blFile, ".")(UBound(Split(blFile, "."))))
	End Function	

	'// 设置文件或文件夹属性
	Public Function [Attributes](ByVal path, ByVal attrType)
		On Error Resume Next
		Dim p,a,i,n,f,at : p = Me.FormatFilePath(path) : n = 0 : [Attributes] = True
		
		If Not ExistsFile(P) Or Not ExistsFolder(P) Then
			[Attributes] = False
			System.Error.Message = "(" & path & ")"
			System.Error.Raise 55
			Exit Function
		End If
		
		If ExistsFile(p) Then
			Set f = FSO.GetFile(p)
		ElseIf ExistsFolder(p) Then
			Set f = FSO.GetFolder(p)
		End If
		at = f.Attributes : a = UCase(attrType)
		If Instr(a,"+")>0 Or Instr(a,"-")>0 Then
			a = System.Text.IIF(Instr(a," ")>0, Split(a," "), Split(a,","))
			For i = 0 To Ubound(a)
				Select Case a(i)
					Case "+R" at = System.Text.IIF(at And 1,at,at+1)
					Case "-R" at = System.Text.IIF(at And 1,at-1,at)
					Case "+H" at = System.Text.IIF(at And 2,at,at+2)
					Case "-H" at = System.Text.IIF(at And 2,at-2,at)
					Case "+S" at = System.Text.IIF(at And 4,at,at+4)
					Case "-S" at = System.Text.IIF(at And 4,at-4,at)
					Case "+A" at = System.Text.IIF(at And 32,at,at+32)
					Case "-A" at = System.Text.IIF(at And 32,at-32,at)
				End Select
			Next
			f.Attributes = at
		Else
			For i = 1 To Len(a)
				Select Case Mid(a,i,1)
					Case "R" n = n + 1
					Case "H" n = n + 2
					Case "S" n = n + 4
				End Select
			Next
			f.Attributes = System.Text.IIF(at And 32,n+32,n)
		End If
		Set f = Nothing
		If Err.Number <> 0 Then
			[Attributes] = False
			System.Error.Message = "(" & path & ")"
			System.Error.Raise 56
		End If
		Err.Clear()
	End Function
	
	'// 获取文件或文件夹信息
	Public Function GetAttributes(ByVal path, ByVal attrType)
		Dim f,s,p : p = Me.FormatFilePath(path)
		If ExistsFile(p) Then
			Set f = FSO.GetFile(p)
		ElseIf ExistsFolder(p) Then
			Set f = FSO.GetFolder(p)
		Else
			GetAttributes = ""
			System.Error.Message = "(" & path & ")"
			System.Error.Raise 57
			Exit Function
		End If
		Select Case LCase(attrType)
			Case "0","name" : s = f.Name
			Case "1","date", "datemodified" : s = f.DateLastModified
			Case "2","datecreated" : s = f.DateCreated
			Case "3","dateaccessed" : s = f.DateLastAccessed
			Case "4","size" : s = FormatSize(f.Size, s_sizeformat)
			Case "5","attr" : s = Attr2Str(f.Attributes)
			Case "6","type" : s = f.Type
			Case Else s = ""
		End Select
		Set f = Nothing
		GetAttributes = s
	End Function	
	
	'// 格式化文件大小
	Public Function FormatSize(Byval fileSize, ByVal level)
		Dim s : s = Int(fileSize) : level = UCase(level)
		FormatSize = System.Text.IIF(s/(1073741824)>0.01,FormatNumber(s/(1073741824),2,-1,0,-1),"0.01") & " GB"
		If s = 0 Then FormatSize = "0 GB"
		If level = "G" Or (level="AUTO" And s>1073741824) Then Exit Function
		FormatSize = System.Text.IIF(s/(1048576)>0.1,FormatNumber(s/(1048576),1,-1,0,-1),"0.1") & " MB"
		If s = 0 Then FormatSize = "0 MB"
		If level = "M" Or (level="AUTO" And s>1048576) Then Exit Function
		FormatSize = System.Text.IIF((s/1024)>1,Int(s/1024),1) & " KB"
		If s = 0 Then FormatSize = "0 KB"
		If Level = "K" Or (level="AUTO" And s>1024) Then Exit Function
		If level = "B" or level = "AUTO" Then
			FormatSize = s & " bytes"
		Else
			FormatSize = s
		End If
	End Function
	
	'// 格式化文件属性
	Private Function Attr2Str(ByVal attrib)
		Dim a,s : a = Int(attrib)
		If a>=2048 Then a = a - 2048
		If a>=1024 Then a = a - 1024
		If a>=32 Then : s = "A" : a = a- 32 : End If
		If a>=16 Then a = a- 16
		If a>=8 Then a = a - 8
		If a>=4 Then : s = "S" & s : a = a- 4 : End If
		If a>=2 Then : s = "H" & s : a = a- 2 : End If
		If a>=1 Then : s = "R" & s : a = a- 1 : End If
		Attr2Str = s
	End Function
	
End Class
%>

<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统Cookie/Application操作类]
'// +--------------------------------------------------------------------------
Class Cls_IO_CAS
	
	'// 声明私有对象
	Private PrAES
	
	'/* 声明公共对象
	
	'// 初始化资源
	Private Sub Class_Initialize()
		PrAES = False
	End Sub
	
	'// 释放资源
	Private Sub Class_Terminate()
	End Sub
	
	Public Property Let AES(ByVal blParam)
		PrAES = System.Text.ToBoolean(blParam)
	End Property
	
	'// 获取一个Cookies值
	Public Function GetCookie(ByVal blParam)
		Dim blCookie, blName, blSubName
		If InStr(blParam, ":") > 0 Then
			blName = System.Text.CLeft(blParam, ":")
			blSubName = System.Text.CRight(blParam, ":")
			If Not System.Text.IsEmptyAndNull(blName) And Not System.Text.IsEmptyAndNull(blSubName) Then
				If Response.Cookies(blName).HasKeys Then blCookie = Request.Cookies(blName)(blSubName)
			End If
		Else
			If Not System.Text.IsEmptyAndNull(blParam) Then blCookie = Request.Cookies(blParam)
		End If
		If Not System.Text.IsEmptyAndNull(blCookie) Then
			If PrAES Then blCookie = System.Security.AES.Decrypt(blCookie)		
			GetCookie = blCookie
		Else GetCookie = "" End If
	End Function
	
	'// 设置一个Cookies值
	Public Sub SetCookie(ByVal blName, ByVal blValue, ByVal blConfig)
		Dim blExpires, blDomain, blPath, blSecure
		If isArray(blConfig) Then
			Dim I: For I = 0 To UBound(blConfig)
				If isDate(blConfig(I)) Then
					blExpires = cDate(blConfig(I))
				ElseIf System.Text.Test(blConfig(I), "INT") Then
					If blConfig(I) <> 0 Then blExpires = Now() + Int(blConfig(I)) / 60 / 24
				ElseIf System.Text.Test(blConfig(I), "DOMAIN") Or System.Text.Test(blConfig(I), "IP") Then
					blDomain = blConfig(I)
				ElseIf InStr(blConfig(I), "/") > 0 Then
					blPath = blConfig(I)
				ElseIf UCase(blConfig(I)) = "TRUE" Or UCase(blConfig(I)) = "FALSE" Then
					blSecure = blConfig(I)
				End If
			Next
		Else
			If isDate(blConfig) Then
				blExpires = cDate(blConfig)
			ElseIf System.Text.Test(blConfig, "INT") Then
				If blConfig <> 0 Then blExpires = Now() + Int(blConfig) / 60 / 24
			ElseIf System.Text.Test(blConfig, "DOMAIN") Or System.Text.Test(blConfig, "IP") Then
				blDomain = blConfig
			ElseIf InStr(blConfig, "/") > 0 Then
				blPath = blConfig
			ElseIf UCase(blConfig) = "TRUE" Or UCase(blConfig) = "FALSE" Then
				blSecure = blConfig
			End If
		End If
		If Not System.Text.IsEmptyAndNull(blValue) Then
			If PrAES Then blValue = System.Security.AES.Encrypt(blValue) End If
		End If
		If InStr(blName, ":") > 0 Then
			Dim blSubName: blSubName = System.Text.CRight(blName, ":")
			blName = System.Text.CLeft(blName, ":")
			Response.Cookies(blName)(blSubName) = blValue
		Else Response.Cookies(blName) = blValue End If
		If Not System.Text.IsEmptyAndNull(blExpires) Then Response.Cookies(blName).Expires = blExpires
		If Not System.Text.IsEmptyAndNull(blDomain) Then Response.Cookies(blName).Domain = blDomain
		If Not System.Text.IsEmptyAndNull(blPath) Then Response.Cookies(blName).Path = blPath
		If Not System.Text.IsEmptyAndNull(blSecure) Then Response.Cookies(blName).Secure = blSecure
	End Sub
	
	'// 删除一个Cookies值
	Public Sub RemoveCookie(ByVal blParam)
		Dim blName, blSubName
		If InStr(blParam, ":") > 0 Then
			blName = System.Text.CLeft(blParam,":")
			blSubName = System.Text.CRight(blParam, ":")
			If Not System.Text.IsEmptyAndNull(blName) And Not System.Text.IsEmptyAndNull(blSubName) Then
				If Response.Cookies(blName).HasKeys Then Response.Cookies(blName)(blSubName) = Empty
			End If
		Else
			If Not System.Text.IsEmptyAndNull(blParam) Then
				Response.Cookies(blParam) = Empty
				Response.Cookies(blParam).Expires = Now()
			End If
		End If
	End Sub
	
	'// 设置一个Application值
	Public Sub SetApplication(ByVal blName, ByRef blData)
		Application.Lock
		If IsObject(blData) Then Set Application(blName) = blData _
		Else Application(blName) = blData
		Application.UnLock
	End Sub
	
	'// 获取一个Application值
	Public Function GetApplication(ByVal blName)
		If Not System.Text.IsEmptyAndNull(blName) Then
			If IsObject(Application(blName)) Then Set GetApplication = Application(blName) _
			Else GetApplication = Application(blName)
		Else GetApplication = Empty End If
	End Function
	
	'// 删除一个Application值
	Public Sub RemoveApplication(ByVal blName)
		Application.Lock
		Application(blName) = Empty
		Application.UnLock
	End Sub	
	
End Class
%>

<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [风声无组件上传类 2.11(http://www.fonshen.com)]
'// +--------------------------------------------------------------------------
Class Cls_Upload

	Private m_TotalSize, m_MaxSize, m_FileType, m_SavePath, m_AutoSave, m_Error, m_Charset
	Private m_dicForm, m_binForm, m_binItem, m_strDate, m_lngTime
	Public	FormItem, FileItem

	Public Property Get Version
		Version = "Fonshen ASP UpLoadClass Version 2.11"
	End Property

	Public Property Get Error
		Error = m_Error
	End Property

	Public Property Get Charset
		Charset = m_Charset
	End Property
	Public Property Let Charset(strCharset)
		m_Charset = strCharset
	End Property

	Public Property Get TotalSize
		TotalSize = m_TotalSize
	End Property
	Public Property Let TotalSize(lngSize)
		If isNumeric(lngSize) Then m_TotalSize = Clng(lngSize)
	End Property

	Public Property Get MaxSize
		MaxSize = m_MaxSize
	End Property
	Public Property Let MaxSize(lngSize)
		If isNumeric(lngSize) Then m_MaxSize = Clng(lngSize)
	End Property

	Public Property Get FileType
		FileType = m_FileType
	End Property
	Public Property Let FileType(strType)
		m_FileType = strType
	End Property

	Public Property Get SavePath
		SavePath = m_SavePath
	End Property
	Public Property Let SavePath(strPath)
		m_SavePath = Replace(strPath, Chr(0), "")
	End Property

	Public Property Get AutoSave
		AutoSave = m_AutoSave
	End Property
	Public Property Let AutoSave(byVal Flag)
		Select Case Flag
			Case 0, 1, 2: m_AutoSave = Flag
		End Select
	End Property

	Private Sub Class_Initialize
		m_Error	   = -1
		m_Charset  = System.Charset
		m_TotalSize= 0
		m_MaxSize  = 153600
		m_FileType = "jpg/gif/png"
		m_SavePath = ""
		m_AutoSave = 0
		Dim dtmNow : dtmNow = Date()
		m_strDate  = Year(dtmNow) & Right("0"&Month(dtmNow), 2) & Right("0" & Day(dtmNow), 2)
		m_lngTime  = Clng(Timer() * 1000)
		Set m_binForm = Server.CreateObject("ADODB.Stream")
		Set m_binItem = Server.CreateObject("ADODB.Stream")
		Set m_dicForm = Server.CreateObject("Scripting.Dictionary")
		m_dicForm.CompareMode = 1
	End Sub

	Private Sub Class_Terminate
		m_dicForm.RemoveAll
		Set m_dicForm = Nothing
		Set m_binItem = Nothing
		Set m_binForm = Nothing
	End Sub

	Public Function Open()
		Open = 0
		If m_Error = -1 Then m_Error = 0 Else Exit Function End If
		Dim lngRequestSize: lngRequestSize = Request.TotalBytes
		If m_TotalSize > 0 And lngRequestSize > m_TotalSize Then
			m_Error = 5: Exit Function
		ElseIf lngRequestSize < 1 Then
			m_Error = 4: Exit Function
		End If

		Dim lngChunkByte: lngChunkByte = 102400
		Dim lngReadSize: lngReadSize = 0
		m_binForm.Type = 1
		m_binForm.Open()
		Do
			m_binForm.Write(Request.BinaryRead(lngChunkByte))
			lngReadSize = lngReadSize + lngChunkByte
			If lngReadSize >= lngRequestSize Then Exit Do
		Loop		
		m_binForm.Position = 0
		Dim binRequestData: binRequestData = m_binForm.Read()

		Dim bCrLf, strSeparator, intSeparator
		bCrLf = ChrB(13) & ChrB(10)
		intSeparator = InstrB(1, binRequestData, bCrLf) - 1
		strSeparator = LeftB(binRequestData, intSeparator)

		Dim strItem, strInam, strFtyp, strPuri, strFnam, strFext, lngFsiz
		Const strSplit = "'"">"
		Dim strFormItem, strFileItem, intTemp, strTemp
		Dim p_start: p_start = intSeparator + 2
		Dim p_end
		Do
			p_end = InStrB(p_start, binRequestData, bCrLf & bCrLf) - 1
			m_binItem.Type = 1
			m_binItem.Open()
			m_binForm.Position = p_start
			m_binForm.CopyTo m_binItem, p_end - p_start
			m_binItem.Position = 0
			m_binItem.Type = 2
			m_binItem.Charset = m_Charset
			strItem = m_binItem.ReadText()
			m_binItem.Close()
			intTemp = Instr(39, strItem, """")
			strInam = Mid(strItem, 39, intTemp - 39)

			p_start = p_end + 4
			p_end = InStrB(p_start, binRequestData, strSeparator) - 1
			m_binItem.Type = 1
			m_binItem.Open()
			m_binForm.Position = p_start
			lngFsiz = p_end - p_start - 2
			m_binForm.CopyTo m_binItem, lngFsiz

			If Instr(intTemp, strItem, "filename=""") <> 0 Then
			If Not m_dicForm.Exists(strInam&"_From") Then
				strFileItem = strFileItem & strSplit & strInam
				If m_binItem.Size <> 0 Then
					intTemp = intTemp + 13
					strFtyp = Mid(strItem, Instr(intTemp, strItem, "Content-Type: ") + 14)
					strPuri = Mid(strItem, intTemp, Instr(intTemp, strItem, """") - intTemp)
					intTemp = InstrRev(strPuri, "\")
					strFnam = Mid(strPuri, intTemp + 1)
					m_dicForm.Add strInam&"_Type", strFtyp
					m_dicForm.Add strInam&"_Name", strFnam
					m_dicForm.Add strInam&"_Path", Left(strPuri, intTemp)
					m_dicForm.Add strInam&"_Size", lngFsiz
					If Instr(strFnam, ".") <> 0 Then strFext = Mid(strFnam, InstrRev(strFnam, ".") + 1) Else strFext = "" End If

					Select Case strFtyp
					Case "image/jpeg", "image/pjpeg", "image/jpg"
						If LCase(strFext) <> "jpg" Then strFext = "jpg"
						m_binItem.Position = 3
						Do While Not m_binItem.EOS
							Do
								intTemp = AscB(m_binItem.Read(1))
							Loop While intTemp = 255 And Not m_binItem.EOS
							
							If intTemp < 192 Or intTemp > 195 Then
								m_binItem.Read(Bin2Val(m_binItem.Read(2)) - 2)
							Else Exit Do End If
							
							Do
								intTemp = AscB(m_binItem.Read(1))
							Loop While intTemp < 255 And Not m_binItem.EOS
						Loop
						m_binItem.Read(3)
						m_dicForm.Add strInam&"_Height", Bin2Val(m_binItem.Read(2))
						m_dicForm.Add strInam&"_Width", Bin2Val(m_binItem.Read(2))
					Case "image/gif"
						If LCase(strFext) <> "gif" Then strFext = "gif"
						m_binItem.Position = 6
						m_dicForm.Add strInam&"_Width", BinVal2(m_binItem.Read(2))
						m_dicForm.Add strInam&"_Height", BinVal2(m_binItem.Read(2))
					Case "image/png"
						If LCase(strFext) <> "png" Then strFext = "png"
						m_binItem.Position = 18
						m_dicForm.Add strInam&"_Width", Bin2Val(m_binItem.Read(2))
						m_binItem.Read(2)
						m_dicForm.Add strInam&"_Height", Bin2Val(m_binItem.Read(2))
					Case "image/bmp"
						If LCase(strFext) <> "bmp" Then strFext = "bmp"
						m_binItem.Position = 18
						m_dicForm.Add strInam&"_Width", BinVal2(m_binItem.Read(4))
						m_dicForm.Add strInam&"_Height", BinVal2(m_binItem.Read(4))
					Case "application/x-shockwave-flash"
						If LCase(strFext) <> "swf" Then strFext = "swf"
						m_binItem.Position = 0
						If Ascb(m_binItem.Read(1)) = 70 Then
							m_binItem.Position = 8
							strTemp = Num2Str(Ascb(m_binItem.Read(1)), 2, 8)
							intTemp = Str2Num(Left(strTemp, 5), 2)
							strTemp = Mid(strTemp, 6)
							While (Len(strTemp) < intTemp * 4)
								strTemp = strTemp & Num2Str(Ascb(m_binItem.Read(1)), 2, 8)
							wend
							m_dicForm.Add strInam&"_Width", Int(Abs(Str2Num(Mid(strTemp, intTemp + 1, intTemp), 2) - Str2Num(Mid(strTemp, 1, intTemp), 2)) / 20)
							m_dicForm.Add strInam&"_Height", Int(Abs(Str2Num(Mid(strTemp, 3 * intTemp + 1, intTemp), 2) - Str2Num(Mid(strTemp, 2 * intTemp + 1, intTemp), 2)) / 20)
						End If
					End Select

					m_dicForm.Add strInam&"_Ext", strFext
					m_dicForm.Add strInam&"_From", p_start
					If m_AutoSave <> 2 Then
						intTemp = GetFerr(lngFsiz, strFext)
						m_dicForm.Add strInam&"_Err", intTemp
						If intTemp = 0 Then
							If m_AutoSave = 0 Then
								strFnam = GetTimeStr()
								If strFext <> "" Then strFnam = strFnam&"."&strFext
							End If
							m_binItem.SaveToFile Server.MapPath(m_SavePath & strFnam), 2
							m_dicForm.Add strInam, strFnam
						End If
					End If
				Else
					m_dicForm.Add strInam & "_Err", -1
				End If
			End If
			Else
				m_binItem.Position = 0
				m_binItem.Type = 2
				m_binItem.Charset = m_Charset
				strTemp = m_binItem.ReadText
				If m_dicForm.Exists(strInam) Then
					m_dicForm(strInam) = m_dicForm(strInam)&","&strTemp
				Else
					strFormItem = strFormItem & strSplit & strInam
					m_dicForm.Add strInam, strTemp
				End If
			End If

			m_binItem.Close()
			p_start = p_end + intSeparator + 2
		Loop Until p_start + 3 > lngRequestSize
		FormItem = Split(strFormItem, strSplit)
		FileItem = Split(strFileItem, strSplit)
		
		Open = lngRequestSize
	End Function

	Private Function GetTimeStr()
		m_lngTime = m_lngTime + 1
		GetTimeStr = m_strDate & Right("00000000"&m_lngTime, 8)
	End Function

	Private Function GetFerr(lngFsiz, strFext)
		Dim intFerr: intFerr = 0
		If lngFsiz > m_MaxSize And m_MaxSize > 0 Then
			If m_Error = 0 Or m_Error = 2 Then m_Error = m_Error + 1
			intFerr = intFerr+1
		End If
		If Instr(1, LCase("/"&m_FileType&"/"), LCase("/"&strFext&"/")) = 0 And m_FileType <> "" Then
			If m_Error < 2 Then m_Error = m_Error + 2
			intFerr = intFerr + 2
		End If
		GetFerr = intFerr
	End Function

	Public Function Save(Item, strFnam)
		Save = False
		If m_dicForm.Exists(Item&"_From") Then
			Dim intFerr, strFext
			strFext = m_dicForm(Item&"_Ext")
			intFerr = GetFerr(m_dicForm(Item&"_Size"), strFext)
			If m_dicForm.Exists(Item&"_Err") Then
				If intFerr = 0 Then m_dicForm(Item&"_Err") = 0 End If
			Else
				m_dicForm.Add Item&"_Err", intFerr
			End If
			If intFerr <> 0 Then Exit Function
			If VarType(strFnam) = 2 Then
				Select Case strFnam
					Case 0:strFnam = GetTimeStr()
						If strFext <> "" Then strFnam = strFnam&"."&strFext
					Case 1:strFnam = m_dicForm(Item&"_Name")
				End Select
			End If
			m_binItem.Type = 1
			m_binItem.Open
			m_binForm.Position = m_dicForm(Item&"_From")
			m_binForm.CopyTo m_binItem, m_dicForm(Item&"_Size")
			m_binItem.SaveToFile Server.MapPath(m_SavePath & strFnam), 2
			m_binItem.Close()
			If m_dicForm.Exists(Item) Then
				m_dicForm(Item) = strFnam
			Else
				m_dicForm.Add Item, strFnam
			End If
			Save = True
		End If
	End Function

	Public Function GetData(Item)
		GetData = ""
		If m_dicForm.Exists(Item&"_From") Then
			If GetFerr(m_dicForm(Item&"_Size"), m_dicForm(Item&"_Ext")) <> 0 Then Exit Function
			m_binForm.Position = m_dicForm(Item&"_From")
			GetData = m_binForm.Read(m_dicForm(Item&"_Size"))
		End If
	End Function

	Public Function Form(Item)
		If m_dicForm.Exists(Item) Then Form = m_dicForm(Item) Else Form = "" End If
	End Function

	Private Function BinVal2(bin)
		Dim lngValue: lngValue = 0
		Dim I: For I = LenB(bin) To 1 Step -1
			lngValue = lngValue *256 + AscB(MidB(bin, I, 1))
		Next
		BinVal2 = lngValue
	End Function

	Private Function Bin2Val(bin)
		Dim lngValue: lngValue = 0
		Dim I: For I = 1 To LenB(bin)
			lngValue = lngValue * 256 + AscB(MidB(bin, I, 1))
		Next
		Bin2Val = lngValue
	End Function

	Private Function Num2Str(num, base, lens)
		Dim I, ret: ret = ""
		While(num >= base)
			I = num Mod base
			ret = I & ret
			num = (num - I) / base
		wend
		Num2Str = Right(String(lens, "0") & num & ret, lens)
	End Function

	Private Function Str2Num(str, base)
		Dim ret: ret = 0 
		Dim I: For I = 1 To Len(str)
			ret = ret * base + Cint(Mid(str, I, 1))
		Next
		Str2Num = ret
	End Function
	
	Public Function Description(ByVal blError)
		Select Case blError
			Case -1: Description = "没有文件上传。"
			Case 0: Description = "上传成功。"
			Case 1: Description = "上传生效，文件大小超过了限制的 " & MaxSize / 1024 & "K，而未被保存。"
			Case 2: Description = "上传生效，文件类型受系统限制，而未被保存。"
			Case 3: Description = "上传生效，文件大小超过了限制的 " & MaxSize / 1024 & "K，且文件类型受系统限制，而未被保存。"
			Case 4: Description = "异常，不存在上传。"
			Case 5: Description = "异常，上传已经取消。"
		End Select
	End Function
	
End Class
%>