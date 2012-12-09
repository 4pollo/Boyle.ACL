<!--#include file="./IO.class.asp"-->
<!--#include file="./Db.class.asp"-->
<!--#include file="./JSON.class.asp"-->
<!--#include file="./Text.class.asp"-->
<!--#include file="./Array.class.asp"-->
<!--#include file="./Cache.class.asp"-->
<!--#include file="./Error.class.asp"-->
<!--#include file="./Security.class.asp"-->
<!--#include file="./Template.class.asp"-->

<!--#include file="./Model.class.asp"-->

<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统接口初始化]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

Class Boyle
	
	'// 定义私有命名对象
	Private PrModel
	Private PrIO, PrData
	Private PrJSON, PrText
	Private PrArray, PrCache, PrError
	Private PrUpload, PrTemplate, PrSecurity
	
	Private PrDebug, PrCharset, PrQueries
	
	'// 定义公共对象
	Public Name, Version
	
	'// 初始化命名对象
	Private Sub Class_Initialize()
		'On Error Resume Next
		'// 定义系统名称和版本
		Name = "Boyle.ACL": Version = "4.0.121028"		
		'// 配置系统所使用的文件编码，系统默认采用UTF-8编码
		PrCharset = "UTF-8"		
		'// 系统默认关闭调试模式
		PrDebug = False
	End Sub
	
	'// 释放命名对象
	Private Sub Class_Terminate()
		TerminateNamespace "IO": TerminateNamespace "Data"
		TerminateNamespace "JSON": TerminateNamespace "Text"
		TerminateNamespace "Array": TerminateNamespace "Cache"
		TerminateNamespace "Error": TerminateNamespace "Upload"
		TerminateNamespace "Template": TerminateNamespace "Security"
		
		TerminateNamespace "Model"
	End Sub
	
	'// 释放命名对象
	Private Sub TerminateNamespace(ByVal blNamespace)
		If IsObject(blNamespace) Then ExecuteGlobal("Set "& blNamespace &" = Nothing") End If
	End Sub
	
	'// 声明模块单元
	Public Property Get Model
		If Not IsObject(PrModel) Then Set PrModel = New Cls_Model End If		
		Set Model = PrModel
	End Property
	Public Property Get IO
		If Not IsObject(PrIO) Then Set PrIO = New Cls_IO End If		
		Set IO = PrIO
	End Property
	Public Property Get Data
		If Not IsObject(PrData) Then Set PrData = New Cls_Data End If		
		Set Data = PrData
	End Property
	Public Property Get Text
		If Not IsObject(PrText) Then Set PrText = New Cls_Text End If		
		Set Text = PrText
	End Property
	Public Property Get JSON
		If Not IsObject(PrJSON) Then Set PrJSON = New Cls_JSON End If		
		Set JSON = PrJSON
	End Property
	Public Property Get [Array]
		If Not IsObject(PrArray) Then Set PrArray = New Cls_Array End If		
		Set [Array] = PrArray
	End Property
	Public Property Get Cache
		If Not IsObject(PrCache) Then Set PrCache = New Cls_Cache End If		
		Set Cache = PrCache
	End Property
	Public Property Get [Error]
		If Not IsObject(PrError) Then Set PrError = New Cls_Error End If		
		Set [Error] = PrError
	End Property
	Public Property Get Upload
		If Not IsObject(PrUpload) Then Set PrUpload = New Cls_Upload End If		
		Set Upload = PrUpload
	End Property
	Public Property Get Template
		If Not IsObject(PrTemplate) Then Set PrTemplate = New Cls_Template End If		
		Set Template = PrTemplate
	End Property
	Public Property Get Security
		If Not IsObject(PrSecurity) Then Set PrSecurity = New Cls_Security End If		
		Set Security = PrSecurity
	End Property
	
	'// 设置是否开启调试功能
	Public Property Let [Debug](ByVal blParam)
		PrDebug = blParam
		[Error].Debug = blParam
	End Property
	Public Property Get [Debug]()
		[Debug] = PrDebug
	End Property
	
	'// 设置系统编码
	Public Property Let Charset(ByVal blParam)
		PrCharset = blParam
		IO.Charset = blParam
	End Property
	Public Property Get Charset()
		Charset = PrCharset
	End Property

	'// 页面执行数据库操作的次数
	Public Property Let Queries(ByVal blNumber)
		PrQueries = PrQueries + blNumber
	End Property
	Public Property Get Queries()
		Queries = PrQueries
	End Property
	
	'// 返回页面执行所用的时间
	Public Property Get [End]()
		[End] = FormatNumber(Timer() - vbTIME, 6, -1)
	End Property
	
	'// 获取地址栏信息
	Public Function Uri(ByVal blParam)
		Dim I, blOut, blItem, blTemp, blHasQueryString, blQueryString
		Dim blScriptName: blScriptName = Request.ServerVariables("SCRIPT_NAME")
		Dim blDir: blDir = Left(blScriptName, InstrRev(blScriptName, "/"))
		Dim blUrl: blUrl = blScriptName: blParam = LCase(blParam)
		
		'// 当值为空时，接收地址栏中的所有数据
		If Text.IsEmptyAndNull(blParam) Then
			Dim blHttp, blPort
			With Request
				If .ServerVariables("HTTPS")="on" Then
					blHttp = "https://"
					blPort = Text.IIF(Int(.ServerVariables("SERVER_PORT")) = 443, "", ":" & .ServerVariables("SERVER_PORT"))
				Else
					blHttp = "http://"
					blPort = Text.IIF(Int(.ServerVariables("SERVER_PORT")) = 80, "", ":" & .ServerVariables("SERVER_PORT"))
				End If
				blUrl = blHttp & .ServerVariables("SERVER_NAME") & blPort & blScriptName
				If Not Text.IsEmptyAndNull(.QueryString()) Then blUrl = blUrl & "?" & .QueryString()
				Uri = blUrl
			End With			
		'// 当值为0时，接收目标文件夹相对路径
		ElseIf blParam = "0" Then Uri = blDir
		'// 当值为1时，接收目标文件相对路径
		ElseIf blParam = "1" Then Uri = blUrl
		Else
			If InStr(blParam, ":") > 0 Then
				blUrl = blDir: blOut = Mid(blParam, 2)
				blHasQueryString = Text.IIF(Text.IsEmptyAndNull(blOut), 0, 1)
			Else blOut = blParam: blHasQueryString = 1 End If
			
			If Not Text.IsEmptyAndNull(Request.QueryString()) Then
				'// 当值为2时，接收目标文件相对路径及所有参数
				If blParam = "2" Or blHasQueryString = 0 Then
					blUrl = blUrl & "?" & Request.QueryString()
				Else
					'// 其它值说明：[NAME] - 接收目标文件相对路径及NAME标签及值
					'//				[-NAME] - 不接收目标文件相对路径及NAME标签及值
					'//				[:] - 接收目标文件夹相对路径的所有标签及值
					'//				[:NAME,NAME1] - 接收目标文件夹相对路径NAME和NAME1标签及其值			
					'//				[:-NAME,-NAME1] - 接收目标文件夹相对路径除NAME和NAME1标签及其值			
					blTemp = "": I = 0: blOut = "," & blOut & ","
					blQueryString = Text.IIF(InStr(blOut, "-") > 0, "Not InStr(blOut, "",-""&blItem&"","") > 0", "InStr(blOut, "",""&blItem&"","") > 0")
					For Each blItem In Request.QueryString()
						If Eval(blQueryString) Then
							If I <> 0 Then blTemp = blTemp & "&"
							blTemp = blTemp & blItem & "=" & Request.QueryString(blItem)
							I = I + 1
						End If
					Next
					If Not Text.IsEmptyAndNull(blTemp) Then blUrl = blUrl & "?" & blTemp
				End If
			End If
			Uri = blUrl
		End If		
	End Function
	
	'// 接收GET方式所传输的数据
	Public Function [Get](ByVal blParam)
		Dim blContent: blParam = Text.Separate(blParam)
		Select Case UCase(blParam(1))
			Case "0", "INT":	'// 当目标标签为空时，值为0
				If Text.IsEmptyAndNull(blParam(0)) Then blContent = 0 _
				Else blContent = Text.ToNumeric(Request.QueryString(blParam(0)))
			Case "", "1", "ALL":'// 当目标标签为空时，接收所有GET数据
				If Text.IsEmptyAndNull(blParam(0)) Then blContent = Request.QueryString() _
				Else blContent = Request.QueryString(blParam(0))
			Case "3", "PAGE":
				blContent = Text.ToNumeric(Request.QueryString(System.Data.Page.Parameters("LABEL")))
			Case Else blContent = blParam(0)
		End Select
		[Get] = Trim(blContent)
	End Function

	'// 接收POST方式所传输的数据
	'// 取Form值，包括上传文件时的普通Form值
	Public Function Post(ByVal strVal)
		Dim blHttpContentType, blFormType
		blHttpContentType = Request.ServerVariables("HTTP_CONTENT_TYPE")		
		If Not Text.IsEmptyAndNull(blHttpContentType) Then blFormType = Split(blHttpContentType, ";")(0) _		
		Else blFormType = "NOUPLOAD" End If
		If LCase(blFormType) = "multipart/form-data" Then
			If Upload.Open() > 0 Then Post = Upload.Form(strVal)
		Else Post = Request.Form(strVal) End If
	End Function
	
	'// 以各种方式输出数据
	Public Sub W(ByVal blParam)
		Response.Write(blParam)
	End Sub
	Public Sub WB(ByVal blParam)
		W blParam & "<br />"
	End Sub
	Public Sub WE(ByVal blParam)
		W blParam: Set System = Nothing: Response.End()
	End Sub

	'// 获取客户端IP地址
	Public Function GetIP()
		Dim addr, x, y
		x = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		y = Request.ServerVariables("REMOTE_ADDR")
		addr = Text.IIF(Text.IsEmptyAndNull(x) Or LCase(x) = "unknown", y, x)
		If InStr(addr, ".") = 0 Then addr = "0.0.0.0"
		GetIP = addr
	End Function

	'// 判断请求是否来自外部
	Public Function IsSelfPost()
		Dim HTTP_REFERER, SERVER_NAME
		HTTP_REFERER = CStr(Request.ServerVariables("HTTP_REFERER"))
		SERVER_NAME  = CStr(Request.ServerVariables("SERVER_NAME"))
		IsSelfPost = False
		IF Mid(HTTP_REFERER, 8, Len(SERVER_NAME)) = SERVER_NAME Then IsSelfPost = True
	End Function
	
End Class

Private vbTIME: vbTIME = Timer()
'// 实例化类，不可更改变量名称
Public System: Set System = New Boyle
%>