<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统调试操作类]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

Class Cls_Error

	'// 定义私有命名对象
	Private PrNumber, PrDelay
	Private PrDebug, PrMessage
	Private PrTitle, PrUrl, PrRedirect
	Private PrError
	
	Private Sub Class_Initialize
		PrNumber    = ""
		PrDelay     = 3000
		PrTitle     = "异常处理"
		PrDebug     = System.Debug
		PrRedirect  = True
		PrUrl       = "javascript:history.go(-1)"
		Set PrError = Server.CreateObject("Scripting.Dictionary")
	End Sub
	Private Sub Class_Terminate
		Set PrError = Nothing
	End Sub
	
	'// 是否开启调试状态（开启后返回开发者错误信息）
	Public Property Get [Debug]
		[Debug] = PrDebug
	End Property
	Public Property Let [Debug](ByVal b)
		PrDebug = b
	End Property
	
	'// 设置或读取自定义的错误代码和错误信息
	Public Default Property Get E(ByVal n)
		If IsNumeric(n) Then n = CLng(n)
		If PrError.Exists(n) Then E = PrError(n) Else E = "未知错误" End If
	End Property
	Public Property Let E(ByVal n, ByVal s)
		If Not System.Text.IsEmptyAndNull(n) And Not System.Text.IsEmptyAndNull(s) Then
			If n > "" Then
				If IsNumeric(n) Then n = CLng(n)
				PrError(n) = s
			End If
		End If
	End Property
	
	'// 取最后一次发生错误的代码
	Public Property Get LastError
		LastError = PrNumber
	End Property
	
	'// 设置和读取错误信息标题
	Public Property Get Title
		Title = PrTitle
	End Property
	Public Property Let Title(ByVal s)
		PrTitle = s
	End Property
	
	'// 设置和读取自定义的附加错误信息
	Public Property Get Message
		Msg = PrMessage
	End Property
	Public Property Let Message(ByVal s)
		PrMessage = s
	End Property
	
	'// 设置和读取页面是否自动转向
	Public Property Get [Redirect]
		[Redirect] = PrRedirect
	End Property
	Public Property Let [Redirect](ByVal b)
		PrRedirect = b
	End Property
	
	'// 设置和读取发生错误后的跳转页地址
	Public Property Get Url
		Url = PrUrl
	End Property
	Public Property Let Url(ByVal s)
		PrUrl = s
	End Property
	
	'// 设置和读取自动跳转页面等待时间（秒）
	Public Property Get Delay
		Delay = PrDelay / 1000
	End Property
	Public Property Let Delay(ByVal i)
		PrDelay = i * 1000
	End Property
	
	'// 生成一个错误
	Public Sub Raise(ByVal n)
		If System.Text.IsEmptyAndNull(n) Then Exit Sub
		PrNumber = n
		If PrDebug Then System.WE ShowMsg(PrError(n) & PrMessage, 1)
		PrMessage = ""
	End Sub
	
	'// 立即抛出一个错误信息
	Public Sub Throw(ByVal blMsg)
		If Left(blMsg, 1) = ":" Then
			blMsg = Mid(blMsg, 2)
			If isNumeric(blMsg) Then blMsg = CLng(blMsg)
			If PrError.Exists(blMsg) Then blMsg = PrError(blMsg)
		End If
		System.W ShowMsg(blMsg, 0)
	End Sub
	
	'// 显示已定义的所有错误代码及信息
	Public Sub Defined()
		Dim key: If Not System.Text.IsEmptyAndNull(PrError) Then
			For Each key In PrError
				System.WB key & " : " & PrError(key)
			Next
		End If
	End Sub
	
	'// 显示错误信息框
	Private Function ShowMsg(ByVal blMsg, ByVal t)
		Dim s, x
		s = s & "<style type=""text/css"">body{margin:50px;font-family: 'Microsoft Yahei', Verdana, arial;font-size:14px;}.dev{color:#999;}"
		s = s & "h2{border-bottom:1px solid #DDD;padding:8px 0;}ul{margin:0;padding:0;list-style:none;}.msg{line-height:200%;}</style>"
		s = s & "<div id=""xError"">" & vbCrLf
		s = s & "<h2>" & PrTitle & "</h2>" & vbCrLf
		s = s & "<p class=""msg"">" & blMsg & "</p>" & vbCrLf		
		If t = 1 Then
			If Err.Number <> 0 Then
				s = s & "<ul class=""dev"">" & vbCrLf
				s = s & "<li>以下信息针对开发者：</li>" & vbCrLf
				s = s & "<li>错误代码：0x" & Hex(Err.Number) & "</li>" & vbCrLf
				s = s & "<li>错误描述：" & Err.Description & "</li>" & vbCrLf
				s = s & "<li>错误来源：" & Err.Source & "</li>" & vbCrLf
				s = s & "</ul>" & vbCrLf
			End If
		Else
			x = System.Text.IIF(PrUrl = "javascript:history.go(-1)", "返回", "继续")
			If PrRedirect Then
				s = s & "<p class=""back"">页面将在" & PrDelay/1000 & "秒钟后跳转，如果浏览器没有正常跳转，<a href=""" & PrUrl & """>请点击此处" & x & "</a></p>" & vbCrLf
				PrUrl = System.Text.IIF(Left(PrUrl, 11) = "javascript:", Mid(PrUrl,12), "location.href='" & PrUrl & "';")
				s = s & System.Text.Format("<{0} type=""text/java{0}"">{1}{2}{3}{1}</{0}>{1}", _
											Array("sc"&"ript", vbCrLf, vbTab , "setTimeout(function(){" & PrUrl & "}," & PrDelay & ");"))
			Else
				s = s & "<p class=""back""><a href=""" & PrUrl & """>请点击此处" & x & "</a></p>" & vbCrLf
			End If
		End If
		s = s & "</div>" & vbCrLf
		ShowMsg = s
	End Function	
End Class
%>