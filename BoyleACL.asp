<!--#include file="./Lib/Core/Boyle.class.asp"-->
<!--#include file="./Common/runtime.asp"-->
<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [框架入口文件]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

'// 导入项目配置文件
System.IO.Import CONF_PATH & "config.asp"

'// 设置调试模式是否开启
System.Debug = C("APP_DEBUG")
'// 设置文件编码
System.Charset = C("DEFAULT_CHARSET")

'// 设置输出的页面编码
Response.Charset = System.Charset
'// 当连接用户断开后自动释放资源
If Not Response.IsClientConnected Then Terminate()

'// 配置数据库连接
System.Data.ConnString = ConfConnString

'// 格式化URL
'// http://localhost/?m=module&a=action&id=1
'// http://localhost/?s=modle/action/var/value
'// http://localhost/app/index.php/Form/read/id/1

'// 根据项目配置的URL访问模式，使用U函数自动获取URL参数
Dim Action, blUri: blUri = U("")
If Not System.Text.IsEmptyAndNull(blUri) Then
	Dim blList: Set blList = System.Array.New
	blList.Symbol = C("URL_PATHINFO_DEPR")
	blList.Data = blUri
	Dim blModel, blAction, blVar, blValue
	Select Case blList.Size
		Case 1:
			blModel = blList(0): blAction = "Index": blVar = "p": blValue = "1"
		Case 2:
			blModel = blList(0): blAction = System.Text.IIF(System.Text.IsEmptyAndNull(blList(1)), "Index", blList(1))
			blVar = "p": blValue = "1"
		Case 3:
			blModel = blList(0): blAction = System.Text.IIF(System.Text.IsEmptyAndNull(blList(1)), "Index", blList(1))
			blVar = blList(2): blValue = "1"
		Case 4:
			blModel = blList(0): blAction = System.Text.IIF(System.Text.IsEmptyAndNull(blList(1)), "Index", blList(1))
			blVar = blList(2): blValue = blList(3)
	End Select
	blList(0) = blModel: blList(1) = blAction
	blList(2) = blVar: blList(3) = blValue

	Call A(blModel, blAction) '// 载入文件
	Set Action = Dicary()
	'On Error Resume Next
	Execute("Set Action("""&blModel&""") = New "&blModel&"Action")
	Execute("Action("""&blModel&""")."&blAction&"("""&blList.J(" ")&""")")
	'If Err Then Response.Redirect("./"): Err.Clear
	Set Action = Nothing: Set blList = Nothing
Else
	Call A("Index", "Index") '// 载入文件
	Set Action = New IndexAction
	Action.Index("Boyle.ACL")
	Set Action = Nothing
End If
%>