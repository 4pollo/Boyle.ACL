<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统JSON操作类]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

'// ---------------------------------------------------------------------------
'// 作者：tugrultopuz@gmail.com
'// 网址：http://code.google.com/p/aspjson/
'// ---------------------------------------------------------------------------

Class Cls_JSON
	
	'// 声明公共对象	
	Public Collection, Count, QuotedVars, StrEncode
	Public Kind '0 = object, 1 = array
	
	Private Sub Class_Initialize
		Set Collection = Dicary()
		
		'// 名称是否用引号
		QuotedVars = True
		'// 是否对中文进行编码
		StrEncode = True
		Count = 0
	End Sub

	Private Sub Class_Terminate
		Set Collection = Nothing
	End Sub
	
	'建新JSON类实例
	Public Function [New](ByVal k)
		Set [New] = New Cls_JSON
		Select Case LCase(k)
			Case "0", "object" [New].Kind = 0
			Case "1", "array"  [New].Kind = 1
		End Select
	End Function

	Private Property Get Counter 
		Counter = Count
		Count = Count + 1
	End Property
	
	'// 设置和读取JSON项的值（值可以是Json对象）
	Public Property Let Pair(p, v)
		If IsNull(p) Then p = Counter
		If varType(v) = 9 Then
			If TypeName(v) = "Cls_JSON" Then Set Collection(p) = v Else Collection(p) = v End If
		Else Collection(p) = v End If
	End Property
	Public Default Property Get Pair(p)
		If IsNull(p) Then p = Count - 1
		If IsObject(Collection(p)) Then Set Pair = Collection(p) Else Pair = Collection(p) End If
	End Property
	
	'// 清除所有JSON项
	Public Sub Clean
		Collection.RemoveAll
	End Sub
	
	'// 删除某一JSON项值
	Public Sub Remove(vProp)
		Collection.Remove vProp
	End Sub
	
	'// 将数据转化Json字符串
	Public Function toJSON(vPair)
		Select Case VarType(vPair)
			Case 0	' Empty
				toJSON = "null"
			Case 1	' Null
				toJSON = "null"
			Case 7	' Date
				toJSON = """" & CStr(vPair) & """"
			Case 8	' String
				toJSON = """" & System.Text.IIF(StrEncode, System.Text.JSEncode(vPair), JSEncode__(vPair)) & """"
			Case 9	' Object
				Dim bFI: bFI = True
				toJSON = toJSON & System.Text.IIF(vPair.Kind, "[", "{")
				Dim I: For Each I In vPair.Collection
					If bFI Then bFI = False Else toJSON = toJSON & ","
					toJSON = toJSON & System.Text.IIF(vPair.Kind, "", System.Text.IIF(QuotedVars, """"&I&"""", I) & ":") & toJSON(vPair(I))
				Next
				toJSON = toJSON & System.Text.IIF(vPair.Kind, "]", "}")
			Case 11
				toJSON = System.Text.IIF(vPair, "true", "false")
			Case 12, 8192, 8204
				toJSON = RenderArray(vPair, 1, "")
			Case Else
				toJSON = Replace(vPair, ",", ".")
		End select
	End Function
	
	'// 递归数组生成Json字符串
	Private Function RenderArray(arr, depth, parent)
		Dim first : first = LBound(arr, depth)
		Dim last : last = UBound(arr, depth)
		Dim index, rendered
		Dim limiter : limiter = ","
		RenderArray = "["
		For index = first To last
			If index = last Then
				limiter = ""
			End If 
			On Error Resume Next
			rendered = RenderArray(arr, depth + 1, parent & index & "," )
			If Err = 9 Then
				On Error GoTo 0
				RenderArray = RenderArray & toJSON(Eval("arr(" & parent & index & ")")) & limiter
			Else
				RenderArray = RenderArray & rendered & "" & limiter
			End If
		Next
		RenderArray = RenderArray & "]": Err.Clear
	End Function
	
	'// 返回Json字符串
	Public Property Get jsString
		jsString = toJSON(Me)
	End Property
	
	'// 输出为Json格式文件
	Public Sub Flush
		Response.Clear()
		Response.Charset = "UTF-8"
		Response.ContentType = "application/json"
		'System.NoCache()
		System.WE jsString
	End Sub
	
	'// 复制JSON对象
	Public Function Clone
		Set Clone = ColClone(Me)
	End Function
	
	Private Function ColClone(blCore)
		Dim DbJSON: Set DbJSON = New Cls_JSON
		DbJSON.Kind = blCore.Kind
		Dim I: For Each I In blCore.Collection
			If IsObject(blCore(I)) Then Set DbJSON(I) = ColClone(blCore(I)) Else DbJSON(I) = blCore(I) End If
		Next
		Set ColClone = DbJSON
	End Function
	
	'// 处理字符串中的Javascript特殊字符，不处理中文
	Private Function JsEncode__(ByVal blParam)
		If System.Text.IsEmptyAndNull(blParam) Then JsEncode__ = "" : Exit Function
		Dim arr1, arr2, I, J, C, P, T
		arr1 = Array(&h27, &h22, &h5C, &h2F, &h08, &h0C, &h0A, &h0D, &h09)
		arr2 = Array(&h27, &h22, &h5C, &h2F, &h62, &h66, &h6E, &h72, &h749)
		For I = 1 To Len(blParam)
			P = True: C = Mid(blParam, I, 1)
			For J = 0 To Ubound(arr1)
				If C = Chr(arr1(J)) Then
					T = T & "\" & Chr(arr2(J))
					P = False: Exit For
				End If
			Next
			If P Then T = T & C
		Next
		JsEncode__ = T
	End Function
	
End Class
%>