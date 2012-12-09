<!--#include file="./common.asp"-->
<!--#include file="../Conf/convention.asp"-->
<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统运行时文件]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

'// 定义系统常量
Private CORE_PATH: CORE_PATH       = BOYLE_PATH & "Lib/Core/"	'// 系统核心类库目录
Private RUNTIME_PATH: RUNTIME_PATH = APP_PATH & "Runtime/"		'// 项目运行时目录
Private LIB_PATH: LIB_PATH         = APP_PATH & "Lib/"			'// 项目类库目录
Private CONF_PATH: CONF_PATH       = APP_PATH & "Conf/"			'// 项目配置目录
Private COMMON_PATH: COMMON_PATH   = APP_PATH & "Common/"		'// 项目公共目录
Private LANG_PATH: LANG_PATH       = APP_PATH & "Lang/"			'// 项目语言目录
Private TMPL_PATH: TMPL_PATH       = APP_PATH & "Tpl/"			'// 项目模板目录
Private HTML_PATH: HTML_PATH       = APP_PATH & "Html/"			'// 项目静态目录
Private CACHE_PATH: CACHE_PATH     = RUNTIME_PATH & "Cache/"	'// 模板缓存目录
Private DATA_PATH: DATA_PATH       = RUNTIME_PATH & "Data/"		'// 数据缓存目录
Private LOG_PATH: LOG_PATH         = RUNTIME_PATH & "Logs/"		'// 日志文件目录
Private TEMP_PATH: TEMP_PATH       = RUNTIME_PATH & "Temp/"		'// 临时缓存目录

'// 加载运行时所需要的文件 并负责自动生成目录
Function load_runtime_file()
	With System.IO
		'// 检查项目目录结构 如果不存在则自动创建
		If Not .ExistsFolder(LIB_PATH) Then
			'// 创建项目目录结构
			build_app_dir()
		End If
	End With
End Function

'// 创建项目目录结构
Function build_app_dir()
	build_app_dir = False
	With System.IO
		Dim I, blDir
		blDir = Array(LIB_PATH, LIB_PATH&"Action/", LIB_PATH&"Model/", LIB_PATH&"Behavior/", LIB_PATH&"Widget/", CONF_PATH, COMMON_PATH, LANG_PATH, TMPL_PATH, RUNTIME_PATH, LOG_PATH, DATA_PATH, TEMP_PATH, CACHE_PATH)
		For I = 0 To Ubound(blDir)
			If Not .ExistsFolder(blDir(I)) Then build_app_dir = .CreateFolder(blDir(I))
		Next
		'// 写入初始配置文件
		If Not .ExistsFile(CONF_PATH&"config.asp") Then .Save CONF_PATH&"config.asp", "<"&"%"&vbCrLf&"'//这个是项目自动生成的配置文件"&vbCrLf&"C(""配置项"") = ""配置值"""&vbCrLf&"%"&">"
		'// 写入测试Action
		If Not .ExistsFile(LIB_PATH&"Action/IndexAction.class.asp") Then build_first_action()
		'// 写入测试模板文件
		If Not .ExistsFile(TMPL_PATH&"Index/index.html") Then build_first_template()
	End With
End Function

'// 创建测试Action
Function build_first_action()
	With System.IO
		build_first_action = .Save(LIB_PATH&"Action/IndexAction.class.asp", .Read(BOYLE_PATH&"Tpl/index_action.tpl"))
	End With
End Function

'// 创建测试模板文件
Function build_first_template()
	With System.IO
		build_first_template = .Save(TMPL_PATH&"Index/index.html", .Read(BOYLE_PATH&"Tpl/index_template.tpl"))
	End With
End Function

'// 生成目录安全文件
Function build_dir_secure(Byval blParam)
	'// 在每个目录下生成一个index.html的空内容的文件
End Function

'// 加载运行时所需文件
load_runtime_file()
%>