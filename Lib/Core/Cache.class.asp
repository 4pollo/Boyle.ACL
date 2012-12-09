<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统缓存操作类]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

'// +--------------------------------------------------------------------------
'// | Coldstone(http://easp.lengshi.com)
'// +--------------------------------------------------------------------------
Class Cls_Cache

	'// 定义公共命名对象
	Public Items, CountEnabled, Expires, FileType
	
	'// 定义私有命名对象
	Private PrPath
	
	'构造函数
	Private Sub Class_Initialize
		Set Items = Server.CreateObject("Scripting.Dictionary")
		PrPath = Server.MapPath("/_cache") & "\"
		CountEnabled = True
		Expires = 5
		FileType = ".cache"
		
		System.Error.E(91) = "当前对象不允许缓存到内存缓存"
		System.Error.E(92) = "缓存文件不存在"
		System.Error.E(93) = "当前内容不允许缓存到文件缓存"
	End Sub
	
	'// 析构函数
	Private Sub Class_Terminate
		Set Items = Nothing
	End Sub
	
	'// 建新缓存类实例
	Public Function [New]()
		Set [New] = New Cls_Cache
	End Function
	
	'// 取当前所有缓存数量
	Public Property Get Count
		Count = System.Text.IIF(CountEnabled, Cls_Cache_Count, -1)
	End Property
	
	'// 添加缓存值
	Public Property Let Item(ByVal p, ByVal v)
		If IsNull(p) Then p = ""
		If Not IsObject(Items(p)) Then
			Set Items(p) = New Cls_Cache_Info
			Items(p).CountEnabled = CountEnabled
			Items(p).Expires = Expires
			Items(p).FileType = FileType
		End If
		Items(p).Name = p
		Items(p).Value = v
		Items(p).SavePath = PrPath
	End Property
	
	'// 获取缓存值
	Public Default Property Get Item(ByVal p)
		If Not IsObject(Items(p)) Then
			Set Items(p) = New Cls_Cache_Info
			Items(p).Name = p
			Items(p).SavePath = PrPath
			Items(p).CountEnabled = CountEnabled
			Items(p).Expires = Expires
			Items(p).FileType = FileType
		End If
		Set Item = Items(p)
	End Property
	
	'// 设置文件缓存保存位置
	Public Property Let SavePath(ByVal blParam)
		If Not Instr(blParam, ":") = 2 Then blParam = Server.MapPath(blParam)
		If Right(blParam,1) <> "\" Then blParam = blParam & "\"
		PrPath = blParam
	End Property
	Public Property Get SavePath()
		SavePath = PrPath
	End Property
	
	'// 保存所有文件缓存
	Public Sub SaveAll
		Dim F: For Each F In Items
			Items(F).Save
		Next
	End Sub
	
	'// 保存所有内存缓存
	Public Sub SaveAppAll  
		Dim F: For Each F In Items
			Items(F).SaveApp
		Next
	End Sub
	
	'// 清除所有文件缓存
	Public Sub RemoveAll
		Dim F: For Each F In Items
			Items(F).Remove
		Next
	End Sub
	
	'// 清除所有内存缓存
	Public Sub RemoveAppAll  
		Dim F: For Each F In Items
			Items(F).RemoveApp
		Next
	End Sub
	
	'// 清空缓存
	Public Sub [Clear]
		RemoveAll: RemoveAppAll
		System.IO.CAS.RemoveApplication "Cls_Cache_Count"
	End Sub
End Class

'// 缓存项处理方法
Class Cls_Cache_Info
	
	'// 定义公共命名对象
	Public SavePath, [Name], CountEnabled, FileType
	
	'// 定义私有命名对象
	Private i_exp, d_exp, o_value
	
	Private Sub Class_Initialize
		i_exp = 5: d_exp = ""
	End Sub
	Private Sub Class_Terminate
		If IsObject(o_value) Then Set o_value = Nothing
	End Sub
	
	'// 设置过期时间
	Public Property Let Expires(ByVal i)
		If isDate(i) Then
			'// 具体日期时间
			d_exp = CDate(i)
		ElseIf isNumeric(i) Then
			'// 数值（分钟）
			If i>0 Then
				i_exp = i
			ElseIf i=0 Then
				i_exp = 60*24*365*99
			End If
		End If
	End Property
	'// 显示过期时间
	Public Property Get Expires()
		Expires = System.Text.IIF(Not System.Text.IsEmptyAndNull(d_exp), d_exp, i_exp)
	End Property
	'// 给当前缓存赋值
	Public Property Let [Value](ByVal blParam)
		If IsObject(blParam) Then
			Select Case TypeName(blParam)
				Case "Recordset"
				'// 如果是记录集
					Set o_value = blParam.Clone
				Case Else
				'// 如果是其它对象
					Set o_value = blParam
			End Select
		Else
			'// 其它直接赋值
			o_value = blParam
		End If
	End Property
	'// 取当前缓存值
	Public Default Property Get [Value]()
		'// 在内存缓存中取值
		Dim app : app = System.IO.CAS.GetApplication(Me.Name)
		If IsArray(app) Then
			If UBound(app) = 1 Then
				If IsDate(app(0)) Then
					If IsObject(app(1)) Then
						Set [Value] = app(1)
						Exit Property
					Else
						[Value] = app(1)
						If Not System.Text.IsEmptyAndNull([Value]) Then Exit Property
					End If
				End If
			End If
		End If
		
		'// 如果内存缓存中没有该值则在文件缓存中取
		If System.IO.ExistsFile(FilePath) Then
			On Error Resume Next
			Dim rs
			Set rs = Server.CreateObject("Adodb.Recordset")
			rs.Open FilePath
			If Err.Number <> 0 Then
				Err.Clear
				[Value] = System.IO.Read(FilePath)
			Else
				Set [Value] = rs
			End If
		Else
			System.Error.Message = "("""&System.Text.HtmlEncode(Me.Name)&""")" : System.Error.Raise 92
		End If
	End Property
	
	'// 保存到内存缓存
	Public Sub SaveApp
		Dim appArr(1) : appArr(0) = Now()
		If IsObject(o_value) Then
			'// 保存字典对象和记录对象（记录集对象将自动转为二维数组）
			Select Case TypeName(o_value)
				Case "Dictionary"
					Set appArr(1) = o_value
				Case "Recordset"
					appArr(1) = o_value.GetRows(-1)
				Case Else
					System.Error.Message = "("""&System.Text.HtmlEncode(Me.Name)&" &gt; "&TypeName(o_value)&""")" : System.Error.Raise 91
			End Select
		Else
			appArr(1) = o_value
		End If
		System.IO.CAS.SetApplication Me.Name, appArr
		If CountEnabled Then Cls_CacheCount_Change Me.Name, 1
	End Sub
	
	'// 保存到文件缓存
	Public Sub Save
		Select Case TypeName(o_value)
			Case "Recordset"
				System.IO.Save FilePath, "rs"
				System.IO.Delete FilePath
				o_value.Save FilePath, 1'adPersistXML
				If CountEnabled Then Cls_CacheCount_Change Me.Name, 1
			Case "String"
				System.IO.Save FilePath, o_value
				If CountEnabled Then Cls_CacheCount_Change Me.Name, 1
			Case Else
				System.Error.Message = "("""&System.Text.HtmlEncode(Me.Name)&""")" : System.Error.Raise 93
		End Select
	End Sub
	
	'// 删除缓存
	Public Sub Remove
		'// 删除文件缓存
		If Not System.Text.Test(DelPath, "[*?]") Then
			If System.IO.ExistsFile(DelPath) Or System.IO.ExistsFolder(DelPath) Then
				If System.IO.ExistsFile(DelPath) Then
					System.IO.Delete(DelPath)
				ElseIf System.IO.ExistsFolder(DelPath) Then
					System.IO.DeleteFolder(DelPath)
				Else
					System.Error.Message = "("& DelPath &")"
					System.Error.Raise "删除失败，文件不存在。"
				End If
			End If
			If CountEnabled Then Cls_CacheCount_Change Me.Name, -1
		Else
			'// 如果有通配符
			System.IO.Delete Left(DelPath, Len(DelPath) - Len(FileType))
			System.IO.DeleteFolder Left(DelPath, Len(DelPath) - Len(FileType))
			If CountEnabled Then Cls_CacheCount_Change Me.Name, -1
		End If
	End Sub
	
	'// 删除内存缓存
	Public Sub RemoveApp
		If Not System.Text.IsEmptyAndNull(Me.Name) Then System.IO.CAS.RemoveApplication Me.Name
		If CountEnabled Then Cls_CacheCount_Change Me.Name, -1
	End Sub
	
	'// 取文件缓存的缓存路径
	Public Property Get FilePath()
		FilePath = TransPath("[\\:""*?<>|\f\n\r\t\v\s]")
	End Property
	
	'// 取文件缓存的缓存地址，可带通配符
	Private Function DelPath()
		DelPath = TransPath("[\\:""<>|\f\n\r\t\v\s]")
	End Function
	
	'// 将名称转换为文件缓存地址
	Private Function TransPath(ByVal fe)
		Dim s_p : s_p = ""
		Dim parr : parr = split(Me.Name,"/")
		Dim i
		for i = 0 to UBound(parr)
			If System.Text.Test(parr(i),fe) Then parr(i)=Server.URLEncode(parr(i))
			s_p = s_p & "_" & parr(i)
			If i < UBound(parr) Then
				s_p = s_p & "\"
			End If
		next
		If s_p="" Then s_p="_"
		TransPath = SavePath & s_p & FileType
	End Function	
	
	'// 缓存是否可用（未过期）
	Public Function Ready()
		Dim app : app = System.IO.CAS.GetApplication(Me.Name)
		Ready = False
		'// 如果是内存缓存
		If IsArray(app) Then
			If UBound(app) = 1 Then
				If IsDate(app(0)) Then
					Ready = isValid(app(0))
					If Ready Then Exit Function
				End If
			End If
		End If
		'// 如果是文件缓存
		If System.IO.ExistsFile(FilePath) Then
			Ready = isValid(System.IO.GetAttributes(FilePath, 1))
		End If
	End Function
	'// 验证时间是否过期
	Private Function isValid(ByVal t)
		If IsDate(t) Then
			If Not System.Text.IsEmptyAndNull(d_exp) Then
				isValid = (DateDiff("s",Now,d_exp) > 0)
			Else
				isValid = (DateDiff("s",t,Now) < i_exp*60)
			End If
		End If
	End Function
End Class

'// 统计缓存数量
Private Function Cls_Cache_Count()
	Cls_Cache_Count = 0
	Dim n : n = System.IO.CAS.GetApplication("Cls_Cache_Count")
	If IsArray(n) Then
		If Ubound(n) = 1 Then Cls_Cache_Count = n(0)
	End If
End Function

'// 缓存计数更改
Private Function Cls_CacheCount_Change(ByVal a, ByVal t)
	Dim n : n = System.IO.CAS.GetApplication("Cls_Cache_Count")
	If isArray(n) Then
		If Ubound(n) = 1 Then
			If TypeName(n(1)) = "Dictionary" Then
				If t = 1 Then n(1)(a) = a
				If t = -1 Then
					If n(1).Exists(a) Then n(1).Remove(a)
				End If
				System.IO.CAS.SetApplication "Cls_Cache_Count", Array(n(1).Count,n(1))
			End If
		End If
	Else
		Dim dic : Set dic = Server.CreateObject("Scripting.Dictionary")
		If t = 1 Then dic(a) = a
		System.IO.CAS.SetApplication "Cls_Cache_Count", Array(System.Text.IIF(t=1, 1, 0), dic)
	End If
End Function
%>