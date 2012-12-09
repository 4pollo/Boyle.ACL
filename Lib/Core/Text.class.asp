<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统字符串操作类]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

Class Cls_Text
	
	'// 定义私有命名对象
	Private PrRegExp

	'// 初始化类
	Private Sub Class_Initialize()
	End Sub
	
	'// 释放类
	Private Sub Class_Terminate()
		If IsObject(PrRegExp) Then Set PrRegExp = Nothing End If
	End Sub
	
	'// 声明正则表达式模块单元
	Public Property Get RegExpX()
		If Not IsObject(PrRegExp) Then
			Set PrRegExp = New RegExp
			PrRegExp.IgnoreCase = True
			PrRegExp.Global = True
		End If
		Set RegExpX = PrRegExp
	End Property
	
	'/**
	' * @功能说明: 	直接返回判断表达式的值
	' * @参数说明:	- blExpression [string] : 判断表达式
	' *				- blParam1 [string] : 成立时返回的值
	' * 			- blParam2 [string] : 不成立时返回的值
	' * @返回值:		- [string] :  字符串
	' */
	Public Function IIF(ByVal blExpression, ByVal blParam1, ByVal blParam2)
		If blExpression Then IIF = blParam1 Else IIF = blParam2 End If
	End Function
	
	'/**
	' * @功能说明: 判断数组/对象/字符串是否为空
	' * @参数说明: - blParam [string] : 源数组/对象/字符串
	' * @返回值:   - [bool] : 布尔值
	' */
	Public Function IsEmptyAndNull(ByVal blParam)
		Dim blReturn: blReturn = False
		Select Case VarType(blParam)
			Case 0, 1: '// vbEmpty, vbNull
				blReturn = True
			Case 8: '// vbString
				If Trim(blParam) = "" Then blReturn = True
			Case 9: '// vbObject
				Select Case TypeName(blParam)
					Case "Nothing", "Empty": blReturn = True
					Case "Recordset":
						If blParam.State = 0 Then blReturn = True
						If blParam.Bof And blParam.Eof Then blReturn = True
					Case "Dictionary":
						If blParam.Count = 0 Then blReturn = True
				End Select
			Case 8192, 8194, 8204, 8209: '// 8192(Array),8204(Variant),8209(Byte)
				If UBound(blParam) = -1 Then blReturn = True
		End Select
		IsEmptyAndNull = blReturn
	End Function
	
	'/**=
	' * @功能说明： 将目标参数转换为布尔值
	' * @参数说明： - blParam [string,int,bool] : 可以用字符或数字来表示
	' * @返回值：   - [bool] : 布尔值
	' */
	Public Function ToBoolean(ByVal blParam)
		ToBoolean = Me.IIF(StrComp(blParam, "True", 1) = 0 Or CBool(ToNumeric(blParam)), True, False)
	End Function
	
	'/**=
	' * @功能说明: 将目标参数转换为数值
	' * @参数说明: - blParam [all] : 需要转换的数据
	' * @返回值:   - [int] : 数值
	' */
	Public Function ToNumeric(ByVal blParam)		
		If IsNumeric(blParam) Then ToNumeric = CDbl(blParam) Else ToNumeric = 0 End If		
	End Function

	'/**
	' * @功能说明: 将目标字符串换为Char型数组
	' * @参数说明: - blParam [string]: 需要转换的字符
	' * @返回值:   - [array] : 数组
	' */
	Public Function ToCharArray(ByVal blParam)
		ReDim CharArray(Len(blParam))
		Dim I: For I = 1 To Len(blParam)
			CharArray(I - 1) = Mid(blParam, I, 1)
		Next
		ToCharArray = CharArray
	End Function
	
	'/**
	' * @功能说明: 转换字符串为HTML实体代码
	' * @参数说明: - blParam [string]: 需要转换的字符串
	' * @返回值:   - [string] : 字符串
	' */
	Public Function ToUniCode(ByVal blParam)
		Dim I, blTmp, blChar, blMid, blAscW
		ToUniCode = Empty: blTmp = Empty
		For I = 1 To Len(blParam)
			blMid = Mid(blParam, I, 1): blAscW = AscW(blMid)
			If blAscW < 0 Then blAscW = blAscW + 65536 End If
			blTmp = blTmp & "&#" & blAscW & ";"
			' If blAscW >= 0 And blAscW <= 128 Then
				' If blChar = "c" Then blTmp = " " & blTmp: blChar = "e" End If
				' blTmp = blTmp & blMid
			' Else
				' If blChar = "e" Then blTmp = blTmp & " ": blChar = "c" End If
				' blTmp = blTmp & "&#" & blAscW & ";"
			' End If
		Next
		ToUniCode = blTmp
	End Function
	
	'/**
	' * @功能说明: 将字符串转换成伪二维数组
	' * @格式参照: [a:1,2,3,4:20,5]
	' * @参数说明: - blParam [string]: 需要转换的数据
	' * 		  - blItemName [array]: 为空时，表示自动为键值命名。为数组时，表示算定义对象的组名称。下标必须和待转换后的数组下标对称
	' * @返回值:   - [array] : 数组
	' */
	Public Function ToArrays(ByVal blParam, ByVal blItemName)
		Dim I, J, blArray1, blArray2, blDic1, blDic2
		Set blDic1 = Server.CreateObject("Scripting.Dictionary")
		If Not Me.IsEmptyAndNull(blParam) Then
			blArray1 = Split(blParam, ":")
			For I = 0 To UBound(blArray1)
				Set blDic2 = Server.CreateObject("Scripting.Dictionary")
				blArray2 = Split(blArray1(I), ",")
				For J = 0 To UBound(blArray2)
					blDic2.Add J, blArray2(J)
				Next
				If Me.IsEmptyAndNull(blItemName) Then blDic1.Add I, blDic2 _
				Else blDic1.Add blItemName(I), blDic2
				Set blDic2 = Nothing
			Next
			Set ToArrays = blDic1: Set blDic1 = Nothing
		Else Set ToArrays = blDic1: Set blDic1 = Nothing End If
	End Function
	
	'// 函数中实现可变参数
	'// 参考http://www.chinacms.org/article.asp?id=224
	Public Function ToHashTable(ByVal blArray)
		Dim blDic: Set blDic = Server.CreateObject("Scripting.Dictionary")
		blDic.CompareMode = 1
		AddToHashTable blDic, blArray
		Set ToHashTable = blDic
		Set blDic = Nothing
	End Function	
	Private Sub AddToHashTable(ByRef blHashObject, ByVal blArray)
		If Me.IsEmptyAndNull(blArray) Then Exit Sub
		Dim I: For I = 0 To UBound(blArray)
			If IsArray(blArray(I)) Then
				If IsObject(blArray(I)(0)) Then
					System.Error.Message = "当前Array(" & I & ")(0)值类型为：" & TypeName(blArray(I)(0)) & " 。"
					System.Error.E(100) = "键不能为对象类型。"
					System.Error.Raise 100
				End If
				If IsObject(blArray(I)(1)) Then Set blHashObject(blArray(I)(0)) = blArray(I)(1) _
				Else blHashObject(blArray(I)(0)) = blArray(I)(1)
			Else
				Dim blString: blString = blArray(I) & ""
				Dim blPos: blPos = InStr(blString, ":")
                If blPos <= 1 Then
					System.Error.Message = blArray(I)
					System.Error.E(101) = "项目不存在，发生在："
					System.Error.Raise 101
				End If
				Dim blName: blName = Me.Separate(blString)(0)
				Dim blValue: blValue = Me.Separate(blString)(1)
				blHashObject(blName) = blValue
			End If
		Next
	End Sub
	
	'// 将Dictionary对象以JSON方式进行输出
	'// blParam参数 name[:totalName][:notjs]
	'// name String (字符串)
	'// 该Json数据在Javascript中的名称
	'// totalName(可选) String (字符串)
	'// 如果不省略此参数，则会在生成的Json字符串中添加一个名称为该参数的表示总记录数的项 
	'// notjs(可选) String (字符串)
	'// 此参数为固定字符串"notjs",如不省略此参数，则输出的Json字符串中不会将中文进行编码 
	Public Function DictionaryToJSON(ByVal blDictionary, ByVal blParam)
		Dim blKey: blKey = blDictionary.Keys
		Dim blItem: blItem = blDictionary.Items
		Dim blCount: blCount =  blDictionary.Count - 1

		Dim blTotal
		Dim blNotJS: blNotJS = False
		Dim blName: blName = Me.Separate(blParam)
		
		If Not Me.IsEmptyAndNull(blName(1)) Then
			blParam = blName(0): blTotal = blName(1)
			blName = Me.Separate(blTotal)
			If Not Me.IsEmptyAndNull(blName(1)) Then
				blTotal = blName(0): blNotJS = (LCase(blName(1)) = "notjs")
			End If
		End If

		Dim blJSON: Set blJSON = System.JSON.New(0)
		If blNotJS Then blJSON.StrEncode = False
		If Not Me.IsEmptyAndNull(blTotal) Then blJSON(blTotal) = blCount
		blJSON(blParam) = System.JSON.New(0)
		Dim I: For I = 0 To blCount
			blJSON(blParam)(blKey(I)) = blItem(I)
		Next
		DictionaryToJSON = blJSON.JsString
		Set blJSON = Nothing
	End Function

	'/**
	' * @功能说明: 生成随机数字或各种形式的字符串
	' * @参数说明: - blParam [int] : 生成的样式
	' * @返回值:   - [string] : 字符串
	' */	
	Public Function Random(ByVal blParam)
		Dim blRandomStrings, blArray, blLength, blTemp, blMatches, blMatch, blStart, blEnd
		blParam = Replace(Replace(Replace(blParam, "\<", Chr(0)), "\>", Chr(1)), "\:", Chr(2))
		blRandomStrings = ""
		If Me.Test(blParam, "(<\d+>|<\d+-\d+>)") Then
			blTemp = blParam: blArray = Separate(blParam)
			If Not Me.IsEmptyAndNull(blArray(1)) Then
				blRandomStrings = blArray(1) : blTemp = blArray(0) : blArray = ""
			End If
			Set blMatches = Me.MatchX(blParam, "(<\d+>|<\d+-\d+>)")
			For Each blMatch In blMatches
				blArray = blMatch.SubMatches(0)
				blLength = Mid(blArray, 2, Len(blArray) - 2)
				If Me.Test(blLength, "^\d+$") Then
					blTemp = Replace(blTemp, blArray, RandomString(blLength, blRandomStrings), 1, 1)
				Else
					blStart = CLeft(blLength, "-")
					blEnd = CRight(blLength, "-")
					blTemp =  Replace(blTemp, blArray, RandomSpaceNumber(blStart, blEnd), 1, 1)
				End If
			Next
			Set blMatches = Nothing
		ElseIf Me.Test(blParam, "^\d+-\d+$") Then
			blStart = CLeft(blParam, "-")
			blEnd = CRight(blParam, "-")
			blTemp = RandomSpaceNumber(blStart, blEnd)
		ElseIf Me.Test(blParam, "^(\d+)|(\d+:.)$") Then
			blLength = blParam: blArray = Separate(blParam)
			If Not Me.IsEmptyAndNull(blArray(1)) Then
				blRandomStrings = blArray(1): blLength = blArray(0): blArray = ""
			End If
			blTemp = RandomString(blLength, blRandomStrings)
		Else blTemp = blParam End If
		Random = Replace(Replace(Replace(blTemp, Chr(0), "<"), Chr(1), ">"), Chr(2), ":")
	End Function
	
	'/**
	' * @功能说明: 随机生成指定长度的字符串
	' * @参数说明: - blLength [int] : 生成的位数
	' * @参数说明: - blAllowString [string] : 指定的字符串
	' * @返回值:   - [int] : 数值
	' */
	Public Function RandomString(ByVal blLength, ByVal blAllowString)		
		If Me.IsEmptyAndNull(blAllowString) Then blAllowString = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
		Dim I: For I = 1 To blLength
			Randomize(Timer) : RandomString = RandomString & Mid(blAllowString, Int(Len(blAllowString) * Rnd + 1), 1)
		Next
	End Function
	
	'/**
	' * @功能说明: 生成X-Y之间的一个随机数
	' * @参数说明: - blMin [int] : 起始数
	' * @参数说明: - blMax [int] : 结束数
	' * @返回值:   - [int] : 数值
	' */
	Public Function RandomSpaceNumber(ByVal blMin, ByVal blMax)
		Randomize(Timer)
		blMin = Me.ToNumeric(blMin): blMax = Me.ToNumeric(blMax)
		RandomSpaceNumber = Int((blMax - blMin + 1) * Rnd + blMin)
	End Function
	
	'/**
	' * @功能说明: 根据正则表达式验证数据合法性
	' * @参数说明: - blParam [string] : 待验证的字符串
	' *    		   - blExpression [string] : 对应标识或自定义正则式
	' * @返回值:   - [bool] : 布尔值
	' */
	Public Function Test(ByVal blParam, ByVal blExpression)
		If Not Me.IsEmptyAndNull(blParam) Then
			Dim blPattern
			Select Case UCase(blExpression)
				Case "IDCARD":	'// 验证是否为合法的身份证号码
					Test = Me.IIF(isIDCard(blParam), True, False) : Exit Function
				Case "ENGLISH":	'// 验证是否只包含英文字母
					blPattern = "^[A-Za-z]+$"
				Case "CHINESE":	'// 验证是否只包含中文字母
					blPattern = "^[\u0391-\uFFE5]+$"
				Case "USERNAME":'// 验证是否是合法的用户名(4-20位，只能是大小写字母及下划线且以字母开头)
					blPattern = "^[a-z]\w{2,19}$"
				Case "EMAIL":	'// 验证是否是合法的邮箱地址
					blPattern = "^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"
				Case "INT":		'// 验证是否为整数
					blPattern = "^[-\+]?\d+$"
				Case "NUMBER":	'// 验证是否为数字
					blPattern = "^\d+$"
				Case "DOUBLE":	'// 验证是否为双精度数字
					blPattern = "^[-\+]?\d+(\.\d+)?$"
				Case "PRICE":	'// 验证是否为价格格式
					blPattern = "^\d+(\.\d+)?$"
				Case "ZIP":		'// 验证是否为合法的邮编
					blPattern = "^[1-9]\d{5}$"
				Case "QQ":		'// 验证是否为合法的QQ号
					blPattern = "^[1-9]\d{4,9}$"
				Case "PHONE":	'// 验证是否为合法的电话号码
					blPattern = "^((\(\d{2,3}\))|(\d{3}\-))?(\(0\d{2,3}\)|0\d{2,3}-)?[1-9]\d{6,7}(\-\d{1,4})?$"
				Case "MOBILE":	'// 验证是否为合法的手机号码
					blPattern = "^((\(\d{2,3}\))|(\d{3}\-))?(1[35][0-9]|189)\d{8}$"
				Case "URL":		'// 验证是否为合法的网址
					blPattern = "^(http|https|ftp):\/\/[A-Za-z0-9]+\.[A-Za-z0-9]+[\/=\?%\-&_~`@[\]\':+!]*([^<>\""])*$"
				Case "DOMAIN":	'// 验证是否为合法域名
					blPattern = "^[A-Za-z0-9\-]+\.([A-Za-z]{2,4}|[A-Za-z]{2,4}\.[A-Za-z]{2})$"
				Case "IP":		'// 验证是否为合法的IP地址
					blPattern = "^(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])$"
				Case Else blPattern = blExpression
			End Select
			RegExpX.Pattern = blPattern
			Test = RegExpX.Test(CStr(blParam))
		Else Test = False End If
	End Function
	
	'/**
	' * @功能说明: 采用正则表达式对目标字符进行替换
	' * @参数说明: - blExpression [string] : 正则表达式
	' *  		   - blParam1 [string] : 待替换的字符串
	' *    		   - blParam2 [string] : 替换成的字符串
	' * @返回值:   - [string] : 字符串
	' */
	Public Function ReplaceX(ByVal blParam1, ByVal blExpression, ByVal blParam2)
		If Not Me.IsEmptyAndNull(blParam1) Then
			RegExpX.Pattern = blExpression
			ReplaceX = RegExpX.Replace(blParam1, blParam2)
		Else ReplaceX = "" End If
	End Function
	Public Function ReplaceXMultiline(ByVal blParam1, ByVal blExpression, ByVal blParam2)
		If Not Me.IsEmptyAndNull(blParam1) Then
			RegExpX.Multiline = True
			RegExpX.Pattern = blExpression
			ReplaceXMultiline = RegExpX.Replace(blParam1, blParam2)
		Else ReplaceXMultiline = "" End If
	End Function
	
	'/**
	' * @功能说明: 对指定的字符串执行正则表达式搜索
	' * @参数说明: - blParam1 [string] : 字符串
	' *  		   - blExpression [string] : 正则表达式
	' * @返回值:   - [matches] : Matches集合
	' */	
	Public Function MatchX(ByVal blParam1, ByVal blExpression)
		RegExpX.Pattern = blExpression
		Set MatchX = RegExpX.Execute(blParam1)
	End Function

	'/**
	' * @功能说明: 得到某字符在目标字符串中出现的次数
	' * @参数说明: - blParam1 [string] : 待查询的字符
	' *  		   - blParam2 [string] : 目标字符
	' *  		   - blParam3 [bool] : 是否区分大小写以及全半角
	' * @返回值:   - [int] : 数值
	' */
	Public Function RepeatTimes(ByVal blParam1, ByVal blParam2, ByVal blParam3)		
		RepeatTimes = Me.ToNumeric(Me.IIF(ToBoolean(blParam3), _
								  (Len(blParam2) - Len(Replace(blParam2, blParam1, "", 1, -1, 1))) / Len(blParam1), _
								  (Len(blParam2) - Len(Replace(blParam2, blParam1, ""))) / Len(blParam1)))
	End Function
	
	'/**
	' * @功能说明: 用数字"0"，在目标数值前进行填充位数
	' * @参数说明: - blParam1 [string]: 需填充字符
	' *			   - blParam2 [int]: 补充的长度
	' * @返回值:   - [string] : 字符串
	' */
	Public Function AppendZero(ByVal blParam1, ByVal blParam2)
		'// 设置最大字符长度不超过20
		If Len(blParam1) <= 20 Then
			Dim I, strZero
			For I = Len(blParam1) To blParam2 - 1
				strZero = strZero & "0"
			Next
		End If
		AppendZero = strZero & blParam1
	End Function
	
	'/**
	' * @功能说明: 计算源字符串长度(一个中文字符为2个字节长)
	' * @参数说明: - blParam1 [string] : 源字符串
	' * @返回值:   - [int] : 数值
	' */
	Public Function Length(ByVal blParam1)
		If Not Me.IsEmptyAndNull(blParam1) Then
			Dim blLength: blLength = 0
			Dim blParam2: blParam2 = Len(Trim(blParam1))
			Dim I: For I = 1 To blParam2
				blLength = Me.IIF(Abs(AscW(Mid(blParam1, I, 1))) > 255, blLength + 2, blLength + 1)
			Next
			Length = ToNumeric(blLength)
		Else Length = 0 End If
	End Function
		
	'/**
	' * @功能说明: 截取字符串左边的指定数量的字符串
	' * @参数说明: - blParam1 [string] : 源字符串
	' * 		   - blParam2 [int:string] : 截取的长度:代替的符号
	' * @返回值:   - [string] : 字符串
	' */
	Function Cut(ByVal blParam1, ByVal blParam2)
		Dim I, blStringLen, blANSI, blSymbol, blSuffix
		blStringLen = Len(blParam1): blANSI = 0: blSymbol = "..."
		blSuffix = Me.Separate(blParam2, ":"): blParam2 = ToNumeric(blSuffix(0))
		If UBound(blSuffix) > 0 Then blSymbol = blSuffix(1)
		For I = 1 to blStringLen
			blANSI = Me.IIF(Abs(AscW(Mid(blParam1, I, 1))) > 255, blANSI + 2, blANSI + 1)
			If blANSI >= blParam2 Then Cut = Left(blParam1, I) & blSymbol: Exit For _
			Else Cut = blParam1 End If
		Next
		Cut = Replace(Cut, Chr(10), "")
	End Function
	
	'/**
	' * @功能说明: 对字符串进行编码，对应JavaScript中的escape()函数
	' * @参数说明: - blParam [string]: 源字符串
	' * @返回值:   - [string] : 字符串
	' */
	Public Function Escape(ByVal blParam)
		If Not Me.IsEmptyAndNull(blParam) Then
			Dim I, blTempString, blANSI, blString: blString = ""
			For I = 1 To Len(blParam)
				blTempString = Mid(blParam,I,1)
				blANSI = AscW(blTempString)
				'// 0-9 A-Z a-z
				If (blANSI >= 48 And blANSI <= 57) Or (blANSI >= 65 And blANSI <= 90) Or (blANSI >= 97 And blANSI <= 122) Then
					blString = blString & blTempString
				ElseIf InStr("@*_+-./", blTempString) > 0 Then
					blString = blString & blTempString
				ElseIf blANSI > 0 And blANSI < 16 Then
					blString = blString & "%0" & Hex(blANSI)
				ElseIf blANSI >= 16 And blANSI < 256 Then
					blString = blString & "%" & Hex(blANSI)
				Else blString = blString & "%u" & Hex(blANSI) End If
			Next
			Escape = blString
		Else Escape = "" End If
	End Function
	
	'/**
	' * @功能说明: 对字符串进行解码，对应JavaScript中的unescape()函数
	' * @参数说明: - blParam [string]: 源字符串
	' * @返回值:   - [string] : 字符串
	' */
	Public Function UnEscape(ByVal blParam)
		Dim blString: blString = ""
		Dim blLocation: blLocation = InStr(blParam, "%")
		Do While blLocation > 0
			blString = blString & Mid(blParam, 1, blLocation - 1)
			If LCase(Mid(blParam, blLocation + 1, 1)) = "u" Then
				blString = blString & ChrW(CLng("&H" & Mid(blParam, blLocation + 2, 4)))
				blParam = Mid(blParam, blLocation + 6)
			Else
				blString = blString & Chr(CLng("&H" & Mid(blParam, blLocation + 1, 2)))
				blParam = Mid(blParam, blLocation + 3)
			End If
			blLocation = InStr(blParam, "%")
		Loop
		UnEscape = blString & blParam
	End Function
	
	'/**
	' * @功能说明: 过滤地址栏参数中的非法字符
	' * @参数说明: - blParam [string]: 源字符串
	' * @返回值:   - [string] : 字符串
	' */
	Public Function Filter(ByVal blParam)
		If Me.IsEmptyAndNull(blParam) Then
			Filter = "": Exit Function
		End If	
		blParam = Replace(blParam, Chr(0), "")
		blParam = Replace(blParam, "'", "‘")
		blParam = Replace(blParam, "%", "％")
		blParam = Replace(blParam, "-", "－")
		blParam = Replace(blParam, " ", "")
		Filter = Trim(blParam)
	End Function

	'/**
	' * @功能说明: 将字符串中的Html代码进行转换
	' * @参数说明: - blParam [string]: 源字符串
	' * @返回值:   - [string] : 字符串
	' */
	Public Function HtmlEnCode(ByVal blParam)
		If Not Me.IsEmptyAndNull(blParam) Then
			blParam = Replace(blParam, Chr(62), "&gt;")   '// >
			blParam = Replace(blParam, Chr(60), "&lt;")   '// <
			blParam = Replace(blParam, Chr(39), "&#39;")  '// '
			blParam = Replace(blParam, Chr(38), "&amp;")  '// &
			blParam = Replace(blParam, Chr(34), "&quot;") '// "
			blParam = Replace(blParam, Chr(32), "&nbsp;") '// 空格
			blParam = Replace(blParam, Chr(13), "")       '// 回车 
			blParam = Replace(blParam, Chr(10), "<br />") '// 换行
			blParam = Replace(blParam, Chr(10)&Chr(10), "<p></p>")
			blParam = Replace(blParam, Chr(9), "&#160;&#160;&#160;&#160;") '// 水平制表符TAB，用160表示以区别空格
			HtmlEnCode = Trim(blParam)
		Else HtmlEnCode = "" End If
	End Function

	'/**
	' * @功能说明: 将字符串中的HTML代码进行转换，对应HtmlEnCode
	' * @参数说明: - blParam [string]: 源字符串
	' * @返回值:   - [string] : 字符串
	' */
	Public Function HtmlDeCode(ByVal blParam)
		If Not Me.IsEmptyAndNull(blParam) Then
			blParam = Replace(blParam, "&gt;", Chr(62))
			blParam = Replace(blParam, "&lt;", Chr(60))
			blParam = Replace(blParam, "&#39;", Chr(39))
			blParam = Replace(blParam, "&amp;", Chr(38))
			blParam = Replace(blParam, "&quot;", Chr(34))
			blParam = Replace(blParam, "&nbsp;", Chr(32))
			blParam = Replace(blParam, "<p></p>", Chr(13))
			blParam = Replace(blParam, "<br />", Chr(10))
			blParam = Replace(blParam, "&#160;&#160;&#160;&#160;", Chr(9))
			HtmlDeCode = Trim(blParam)
		Else HtmlDeCode = "" End If
	End Function

	'// 功能说明: 处理字符串中的JavaScript特殊字符
	Public Function JSEncode(ByVal blParam)
		If Me.IsEmptyAndNull(blParam) Then JSEncode = "": Exit Function
		
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
			'// 处理中文字符
			If P Then 
				Dim A: A = AscW(C)
				If A > 31 And A < 127 Then
					T = T & C
				ElseIf A > -1 Or A < 65535 Then
					T = T & "\u" & String(4 - Len(Hex(A)), "0") & Hex(A)
				End If 
			End If
		Next
		JSEncode = T
	End Function

	'/**
	' * @功能说明: 清除字符串中的HTML代码
	' * @参数说明: - blParam [string]: 源字符串
	' * @返回值:   - [string] : 字符串
	' */
	Public Function RemoveHtml(ByVal blParam)
		If Not Me.IsEmptyAndNull(blParam) Then
			blParam = ReplaceX(blParam, "<[^>]+>|</[^>]+>", "")
			blParam = Replace(blParam, "<", "&lt;")
			blParam = Replace(blParam, ">", "&gt;")
			RemoveHtml = Trim(blParam)
		Else RemoveHtml = "" End If
	End Function
	
	'/**
	' * @功能说明: 过滤SQL语句中的非法字符
	' * @参数说明: - blParam [string]: 源字符串
	' * @返回值:   - [string] : 字符串
	' */
	Public Function FormatSQLInput(ByVal blParam)
		If Not Me.IsEmptyAndNull(blParam) Then
			blParam = Replace(blParam, "<", "&lt;")
			blParam = Replace(blParam, ">", "&gt;")
			blParam = Replace(blParam, "[", "&#091;")
			blParam = Replace(blParam, "]", "&#093;")
			blParam = Replace(blParam, """", "", 1, -1, 1)
			blParam = Replace(blParam, "=", "&#061;", 1, -1, 1)
			blParam = Replace(blParam, "'", "''", 1, -1, 1)
			blParam = Replace(blParam, "join", "jo&#105;n", 1, -1, 1)
			blParam = Replace(blParam, "like", "lik&#101;", 1, -1, 1)
			blParam = Replace(blParam, "drop", "dro&#112;", 1, -1, 1)
			blParam = Replace(blParam, "cast", "ca&#115;t", 1, -1, 1)	
			blParam = Replace(blParam, "alter", "alt&#101;r", 1, -1, 1)
			blParam = Replace(blParam, "union", "un&#105;on", 1, -1, 1)
			blParam = Replace(blParam, "where", "wh&#101;re", 1, -1, 1)
			blParam = Replace(blParam, "select", "sel&#101;ct", 1, -1, 1)
			blParam = Replace(blParam, "insert", "ins&#101;rt", 1, -1, 1)
			blParam = Replace(blParam, "delete", "del&#101;te", 1, -1, 1)
			blParam = Replace(blParam, "update", "up&#100;ate", 1, -1, 1)
			blParam = Replace(blParam, "create", "cr&#101;ate", 1, -1, 1)
			blParam = Replace(blParam, "modify", "mod&#105;fy", 1, -1, 1)
			blParam = Replace(blParam, "rename", "ren&#097;me", 1, -1, 1)
			FormatSQLInput = Trim(blParam)
		Else FormatSQLInput = "" End If
	End Function
	
	'/**
	' * @功能说明: 转换日文字符为Unicode编码
	' * @参数说明: - blParam [string]: 源字符串
	' * @返回值:   - [string] : 字符串
	' */
	Public Function EnCodeJapanese(ByVal blParam)
		If Not Me.IsEmptyAndNull(blParam) Then
			blParam = Replace(blParam, "ガ", "&#12460;")
			blParam = Replace(blParam, "ギ", "&#12462;")
			blParam = Replace(blParam, "グ", "&#12464;")
			blParam = Replace(blParam, "ア", "&#12450;")
			blParam = Replace(blParam, "ゲ", "&#12466;")
			blParam = Replace(blParam, "ゴ", "&#12468;")
			blParam = Replace(blParam, "ザ", "&#12470;")
			blParam = Replace(blParam, "ジ", "&#12472;")
			blParam = Replace(blParam, "ズ", "&#12474;")
			blParam = Replace(blParam, "ゼ", "&#12476;")
			blParam = Replace(blParam, "ゾ", "&#12478;")
			blParam = Replace(blParam, "ダ", "&#12480;")
			blParam = Replace(blParam, "ヂ", "&#12482;")
			blParam = Replace(blParam, "ヅ", "&#12485;")
			blParam = Replace(blParam, "デ", "&#12487;")
			blParam = Replace(blParam, "ド", "&#12489;")
			blParam = Replace(blParam, "バ", "&#12496;")
			blParam = Replace(blParam, "パ", "&#12497;")
			blParam = Replace(blParam, "ビ", "&#12499;")
			blParam = Replace(blParam, "ピ", "&#12500;")
			blParam = Replace(blParam, "ブ", "&#12502;")
			blParam = Replace(blParam, "ブ", "&#12502;")
			blParam = Replace(blParam, "プ", "&#12503;")
			blParam = Replace(blParam, "ベ", "&#12505;")
			blParam = Replace(blParam, "ペ", "&#12506;")
			blParam = Replace(blParam, "ボ", "&#12508;")
			blParam = Replace(blParam, "ポ", "&#12509;")
			blParam = Replace(blParam, "ヴ", "&#12532;")			
			EnCodeJapanese = Trim(blParam)
		Else EnCodeJapanese = "" End If
	End Function

	'/**
	' * @功能说明: 格式化时间格式
	' * @参数说明: - blDateTime [date]: 日期时间
	' *  		   - blShowType [string]: 格式化的类型
	' * @返回值:   - [string]: 字符串
	' */
	Public Function FormatTime(ByVal blDateTime, ByVal blShowType)
		Dim DateMonth, DateDay, DateHour, DateMinute, DateWeek, DateSecond, DateAMPM
		Dim FullWeekday, ShortWeekday, FullMonth, ShortMonth
		
		FullWeekday = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
		ShortWeekday = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
		FullMonth = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
		ShortMonth = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
		
		DateMonth = Month(blDateTime): DateWeek = Weekday(blDateTime): DateDay = Day(blDateTime)
		DateHour = Hour(blDateTime): DateMinute = Minute(blDateTime): DateSecond = Second(blDateTime)
		If Len(DateMonth) < 2 Then DateMonth = "0" & DateMonth
		If Len(DateDay) < 2 Then DateDay = "0" & DateDay
		If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
		
		Select Case blShowType
			Case "d m y H:I:S A"
				If DateHour > 12 Then DateHour = DateHour - 12: DateAMPM = "PM" _
				Else DateHour = DateHour: DateAMPM = "AM"
				If Len(DateHour) < 2 Then DateHour = "0" & DateHour
				If Len(DateSecond) < 2 Then DateSecond = "0" & DateSecond
				FormatTime = DateDay & " " & Left(FullMonth(DateMonth - 1), 3) & " " & Right(Year(blDateTime), 4) & " " & DateHour & ":" & DateMinute & ":" & DateSecond & " " & DateAMPM
			Case "Y-m-d":
				FormatTime = Year(blDateTime) & "-" & DateMonth & "-" & DateDay
			Case "Y-m-d H:I A":
				If DateHour > 12 Then DateHour = DateHour - 12: DateAMPM = "PM" _
				Else DateHour = DateHour: DateAMPM = "AM"
				If Len(DateHour) < 2 Then DateHour = "0" & DateHour
				If Len(DateMinute) < 2 Then DateMinute = "0" & DateMinute
				FormatTime = Year(blDateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute & " " & DateAMPM
			Case "Y-m-d H:I:S":
				If Len(DateHour) < 2 Then DateHour = "0" & DateHour
				If Len(DateSecond) < 2 Then DateSecond = "0" & DateSecond
				FormatTime = Year(blDateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute & ":" & DateSecond
			Case "YmdHIS":
				DateSecond = Second(blDateTime)
				If Len(DateHour) < 2 Then DateHour = "0" & DateHour
				If Len(DateSecond) < 2 Then DateSecond = "0" & DateSecond
				FormatTime = Year(blDateTime) & DateMonth & DateDay & DateHour & DateMinute & DateSecond
			Case "ym":
				FormatTime = Right(Year(blDateTime), 2) & DateMonth
			Case "d":
				FormatTime = DateDay
			Case "ymd":
				FormatTime = Right(Year(blDateTime), 4) & DateMonth & DateDay
			Case "mdy":
				Dim DayEnd
				Select Case DateDay
					Case 1: DayEnd = "st"
					Case 2: DayEnd = "nd"
					Case 3: DayEnd = "rd"
					Case Else DayEnd = "th"
				End Select
				FormatTime = FullMonth(DateMonth - 1) & " " & DateDay & DayEnd & " " & Right(Year(blDateTime), 4)
			Case "w,d m y H:I:S":
				DateSecond = Second(blDateTime)
				If Len(DateHour) < 2 Then DateHour = "0" & DateHour
				If Len(DateSecond) < 2 Then DateSecond = "0" & DateSecond
				FormatTime = ShortWeekday(DateWeek - 1) & "," & DateDay & " " & Left(FullMonth(DateMonth - 1), 3) & " " & Right(Year(blDateTime), 4) & " " & DateHour & ":" & DateMinute & ":" & DateSecond & " +0800"
			Case "y-m-dTH:I:S":
				If Len(DateHour) < 2 Then DateHour = "0" & DateHour
				If Len(DateSecond) < 2 Then DateSecond = "0" & DateSecond
				FormatTime = Year(blDateTime) & "-" & DateMonth & "-" & DateDay & "T" & DateHour & ":" & DateMinute & ":" & DateSecond & " +08:00"
			Case Else
				If Len(DateHour) < 2 Then DateHour = "0" & DateHour
				FormatTime = Year(blDateTime) & "-" & DateMonth & "-" & DateDay & " " & DateHour & ":" & DateMinute
		End Select
	End Function
	
	'/**
	' * @功能说明: 验证身份证号码是否有效
	' * @参数说明: - blParam [string]: 身份证号码
	' * @返回值:   - [bool]: 布尔值
	' */	
	Private Function isIDCard(ByVal blParam)
		Dim Ai, BirthDay, arrVerifyCode, Wi, I, AiPlusWi, modValue, strVerifyCode
		isIDCard = False
		If Len(blParam) <> 15 And Len(blParam) <> 18 Then Exit Function
		Ai = Me.IIF(Len(blParam) = 18,Mid(blParam, 1, 17),Left(blParam, 6) & "19" & Mid(blParam, 7, 9))
		If Not IsNumeric(Ai) Then Exit Function
		If Not Test(Left(Ai,6), "^(1[1-5]|2[1-3]|3[1-7]|4[1-6]|5[0-4]|6[1-5]|8[12]|91)\d{2}[01238]\d{1}$") Then Exit Function
		BirthDay = Mid(Ai, 7, 4) & "-" & Mid(Ai, 11, 2) & "-" & Mid(Ai, 13, 2)
		If IsDate(BirthDay) Then
			If cDate(BirthDay) > Date() Or cDate(BirthDay) < cDate("1870-1-1") Then Exit Function
		Else Exit Function End If
		arrVerifyCode = Split("1,0,x,9,8,7,6,5,4,3,2", ",")
		Wi = Split("7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2", ",")
		For I = 0 To 16
			AiPlusWi = AiPlusWi + CInt(Mid(Ai, I + 1, 1)) * Wi(I)
		Next
		modValue = AiPlusWi Mod 11
		strVerifyCode = arrVerifyCode(modValue)
		Ai = Ai & strVerifyCode
		If Len(blParam) = 18 And LCase(blParam) <> Ai Then Exit Function
		isIDCard = True
	End Function
	
	'/**
	' * @功能说明: 将复杂的各类集合对象格式化为字符串
	' * @参数说明: - blString [string]: 源字符串
	' * 		   - blValue [all]: 集合对象
	' * @返回值:   - [string]: 字符串
	' */
	Public Function Format(ByVal blString, ByVal blValue)
		Format = FormatString(blString, blValue, 0)
	End Function
	
	'/**
	' * @功能说明: 将复杂的各类集合对象格式化为字符串
	' * @参数说明: - blString [string]: 源字符串
	' * 		   - blValue [all]: 集合对象
	' * 		   - blIndex [int]: 起始索引
	' * @返回值:   - [string]: 字符串
	' */
	Private Function FormatString(ByVal blString, ByRef blValue, ByVal blIndex)
		Dim I, K
		blString = Replace(blString, "\\", Chr(0))
		blString = Replace(blString, "\{", Chr(1))
		Select Case VarType(blValue)
			Case 8192, 8194, 8204, 8209: '// vbArray
				For I = 0 To Ubound(blValue)
					blString = FormatReplace(blString, I + blIndex, blValue(I))
				Next
			Case 9: '// vbObject
				Select Case TypeName(blValue)
					Case "Recordset":
						For I = 0 To blValue.Fields.Count - 1
							blString = FormatReplace(blString, I + blIndex, blValue(I))
							blString = FormatReplace(blString, blValue.Fields.Item(I + blIndex).Name, blValue(I))
						Next
					Case "Dictionary":
						For Each K In blValue
							blString = FormatReplace(blString, K, blValue(K))
						Next
					Case "ISubMatches", "SubMatches":
						For I = 0 To blValue.Count - 1
							blString = FormatReplace(blString, I + blIndex, blValue(I))
						Next
				End Select
			Case 8: '// vbString
				Select Case TypeName(blValue)
					Case "IMatch2", "blMatch":
						blString = FormatReplace(blString, blIndex, blValue.Value)
						For I = 0 To blValue.SubMatches.Count - 1
							blString = FormatReplace(blString, I + blIndex + 1, blValue.SubMatches(I))
						Next
					Case Else blString = FormatReplace(blString, blIndex, blValue)
				End Select
			Case Else blString = FormatReplace(blString, blIndex, blValue)
		End Select
		blString = Replace(blString, Chr(1), "{")
		FormatString = Replace(blString, Chr(0), "\")
	End Function
	
	'// 替换内容
	Private Function FormatReplace(ByVal blString, ByVal blIndex, ByVal blValue)
		Dim blTemp, blRule, blContent, blKind, blMatches, blMatch
		blValue = Me.IIF(Not Me.IsEmptyAndNull(blValue), blValue, "")
		blRule = "\{" & blIndex & "(:((N[,\(%]?(\d+)?)|(D[^\}]+)|(E[^\}]+)|U|L))\}"
		If Me.Test(blString, blRule) Then
			Set blMatches = Me.MatchX(blString, blRule)
			For Each blMatch In blMatches
				blKind = Me.ReplaceX(blMatch.Value, blRule, "$2")
				blContent = "{" & blIndex & ":" & blKind & "}"
				Select Case Left(blKind, 1)
					Case "N":
						If isNumeric(blValue) Then
							Dim blFormat, blGroup, blParens, blPercent, blDecimal
							blFormat = Me.ReplaceX(blKind, "N([,\(%])?(\d+)?", "$1")
							If blFormat = "," Then blGroup = -1
							If blFormat = "(" Then blParens = -1
							If blFormat = "%" Then blPercent = -1
							blDecimal = Me.ReplaceX(blKind, "N([,\(%])?(\d+)?", "$2")
							'// 当N后面不跟参数时，直接输出目标原数值
							If Me.IsEmptyAndNull(blFormat) And Me.IsEmptyAndNull(blDecimal) Then
								blString = Replace(blString, blContent, blValue, 1, -1, 1)
							Else
								blDecimal = Me.IIF(Not Me.IsEmptyAndNull(blDecimal), blDecimal, -1)
								If blPercent Then blString = Replace(blString, blContent, FormatNumber(blValue * 100, blDecimal, -1)&"%", 1, -1, 1) _
								Else blString = Replace(blString, blContent, FormatNumber(blValue, blDecimal, -1, blParens, blGroup), 1, -1, 1)
							End If
						End If
					Case "D": If isDate(blValue) Then blString = Replace(blString, blContent, Me.FormatTime(blValue, Mid(blKind, 2)), 1, -1, 1)
					Case "U": blString = Replace(blString, blContent, UCase(blValue), 1, -1, 1)
					Case "L": blString = Replace(blString, blContent, LCase(blValue), 1, -1, 1)
					Case "E":
						blTemp = Replace(Mid(blKind, 2), "%s", "blValue")
						blTemp = Eval(blTemp)
						blString = Replace(blString, blContent, blTemp, 1, -1, 1)
				End Select
			Next
		Else blString = Replace(blString, "{" & blIndex & "}", blValue, 1, -1, 1) End If
		FormatReplace = blString
	End Function
	
	'/**
	' * @功能说明: 将目标字符串以":"在其第一次出现的位置分隔开
	' * @参数说明: - blParam [string]: 源字符串
	' * @返回值:   - [array] : 数组
	' */
	Public Function Separate(ByVal blParam)
		Dim blArray(1): Dim blPosition: blPosition = InStr(blParam, ":")
		If blPosition > 0 Then blArray(0) = Left(blParam, blPosition - 1): blArray(1) = Mid(blParam, blPosition + 1) _
		Else blArray(0) = blParam: blArray(1) = ""
		Separate = blArray
	End Function
	
	'/**
	' * @功能说明: 截取用某个特殊字符分隔的字符串的特殊字符左边部分
	' * @参数说明: - blString [string]: 源字符串
	' * 		   - blSymbol [string]: 分隔符号
	' * @返回值:   - [string] : 字符串
	' */
	Public Function CLeft(ByVal blString, ByVal blSymbol)
		Dim blPosition: blPosition = InStr(blString, blSymbol)
		If blPosition > 0 Then CLeft = Left(blString, blPosition - 1) _
		Else CLeft = blString
	End Function
	
	'/**
	' * @功能说明: 截取用某个特殊字符分隔的字符串的特殊字符右边部分
	' * @参数说明: - blString [string]: 源字符串
	' * 		   - blSymbol [string]: 分隔符号
	' * @返回值:   - [array] : 字符串
	' */
	Public Function CRight(ByVal blString, ByVal blSymbol)
		Dim blPosition: blPosition = InStr(blString, blSymbol)
		If blPosition > 0 Then CRight = Mid(blString, blPosition + Len(blSymbol)) _
		Else CRight = blString
	End Function
	
End Class
%>