<%
'// 本类由系统自动生成，仅供测试用途
Class IndexAction

	Private Sub Class_Initialize
	End Sub
	
	'// 释放资源
	Private Sub Class_Terminate()
		Call Terminate()
	End Sub

	'// 此方法为系统默认，请不要删除
	Public Sub Index(ByVal blParam)
		System.WB "<a href=""?m=index&a=parts"">普通访问模式</a>&nbsp;<a href=""?"&C("VAR_PATHINFO")&"=index"&C("URL_PATHINFO_DEPR")&"parts"">单参数访问模式</a>"
	End Sub

	Public Sub Parts(ByVal blUrlParam)
		With System.Template
			.d("title") = "Boyle.ACL 示例"

			'// 获取地址栏数据并转换为数组
			blUrlParam = System.Array.NewArray(blUrlParam).Data
			'// 获取值，根据URL访问模式，自动获取值
			Dim blPage
			If C("URL_MODEL") = 0 Then blPage = System.Get(":PAGE")
			If C("URL_MODEL") = 1 Then blPage = blUrlParam(3)
			If C("URL_MODEL") = 2 Then blPage = ""

			Dim Parts: Set Parts = M("PARTS")
			Parts.Parameters("") = Array("CURRENTPAGE:"&blPage&"", "FIELD:ID,CP_NAME,CP_LOCALITY,CP_CAR")
			Parts.Parameters("URL") = U(blUrlParam)
			Dim PagerResult: PagerResult = Parts.Pager()
			.d("parts") = Array(PagerResult(0), "id,name,locality,car")
			.d("pager") = PagerResult(1)
			.d("sql") = PagerResult(2)("SQL")
			Set Parts = Nothing

			.Display()
		End With
	End Sub
End Class
%>