<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统基础函数库]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

'// 获取和设置配置参数
Public C: Set C = Dicary():C.CompareMode = 1

'// 创建一个字典对象
Public Function Dicary()
	Set Dicary = Server.CreateObject("Scripting.Dictionary")
End Function

'// 释放数据源和基类
Public Sub Terminate()
	System.Data.DisConnect()
	Set System = Nothing
	Set C = Nothing
End Sub

'// 配置数据库连接字符串
Public Function ConfConnString()
	If Not System.Text.IsEmptyAndNull(C("DB_NAME")) Then
		Dim TempStr
		Select Case UCase(C("DB_TYPE"))
			Case "0", "MSSQL":
				TempStr = "Provider=sqloledb;Data Source=" & C("DB_HOST") & ";Initial Catalog="& C("DB_NAME") &";User Id="& C("DB_USER") &";Password="& C("DB_PWD") &";"
			Case "1", "ACCESS":
				TempStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & System.IO.FormatFilePath(C("DB_NAME")) & ";Jet OLEDB:Database Password="&C("DB_PWD")&";"
			Case "2", "MYSQL":
				TempStr = "Driver={mySQL};Server="& C("DB_HOST") &";Port="& C("DB_PORT") &";Option=131072;Stmt=;Database="& C("DB_NAME") &";Uid="& C("DB_USER") &";Pwd="& C("DB_PWD") &";"
			Case "3", "ORACLE":
				TempStr = "Provider=msdaora;Data Source="& C("DB_NAME") &";User Id="& C("DB_USER") &";Password="& C("DB_PWD") &";"
		End Select
		ConfConnString = TempStr
	End If
End Function

'// D函数用于实例化Model 格式 项目://分组/模块
'Public Function D()
'End Function

'// M函数用于实例化模型类
Public Function M(byVal blParam)
	Set M = System.Model.New(C("DB_PREFIX") & blParam)
End Function

'// A函数用于实例化动作类
Private Sub A(ByVal blModel, ByVal blAction)
	'// 读取类文件
	System.IO.Import(LIB_PATH & "Action/"&blModel&"Action.class.asp")
	'// 设置模板路径及文件
	System.Template.Root = TMPL_PATH&blModel&"/"
	System.Template.File = blAction
End Sub

'// URL组装 支持不同URL模式
'// 对获取URL参数进行安全过滤，请在这里进行设置
Public Function U(ByVal blParam)
	'// 当参数为空时，则获取URL值
	If System.Text.IsEmptyAndNull(blParam) Then
		If C("URL_MODEL") = 0 Then
			Dim blUrlModel: blUrlModel = LCase(System.Get(C("VAR_MODULE")))
			Dim blUrlAction: blUrlAction = LCase(System.Get(C("VAR_ACTION")))
			If Not System.Text.IsEmptyAndNull(blUrlModel) Then
				U = blUrlModel&C("URL_PATHINFO_DEPR")&blUrlAction
			End If
		ElseIf C("URL_MODEL") = 1 Then
			U = LCase(System.Get(C("VAR_PATHINFO")))
		ElseIf C("URL_MODEL") = 2 Then
			U = ""
		End If
	Else 							'// 否则根据URL访问模式生成相对应的URL地址
		If C("URL_MODEL") = 0 Then	'// 普通模式下，不需要设置URL地址，系统会自动获取
			U = ""
		ElseIf C("URL_MODEL") = 1 Then
			Dim blUrlParam: blUrlParam = System.Array.NewArray(blParam).Data
			U = "?"&C("VAR_PATHINFO")&"="&blUrlParam(0)&C("URL_PATHINFO_DEPR")&blUrlParam(1)&C("URL_PATHINFO_DEPR")&blUrlParam(2)&C("URL_PATHINFO_DEPR")&"*"
		ElseIf C("URL_MODEL") = 2 Then
		End If
	End If
End Function

'// 缓存管理
Public Function S()
End Function

'// 快速文件数据读取和保存 针对简单类型数据 字符串、数组
Public Function F()
End Function

'// 获取和设置语言定义(不区分大小写)
Public Function L()
End Function
%>