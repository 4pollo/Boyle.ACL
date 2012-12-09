<%
'// +----------------------------------------------------------------------
'// | Boyle.ACL [系统模型操作类]
'// +----------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +----------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +----------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +----------------------------------------------------------------------

Class Cls_Model

	'// 定义私有命名对象
	Private PrDic
	Private PrTable

	Private Sub Class_Initialize
		Set PrDic = Dicary(): PrDic.CompareMode = 1
		
		PrTable = "" '// 初始化表格名称
	End Sub
	Private Sub Class_Terminate
		Set PrDic = Nothing
	End Sub
	
	'// 新建类实例
	Public Function [New](ByVal bParam)
		Set [New] = New Cls_Model
		[New].Table = bParam
	End Function

	'// 设置读取表格名称
	Public Property Get Table()
		Table = PrTable
	End Property
	Public Property Let Table(ByVal bParam)
		PrTable = bParam
	End Property

	'// 实现批量设置SQL语句参数
	Public Property Let Parameters(ByVal bField, ByVal bValue)
		'// 当键值为空时，表示对所有参数进行设置
		If System.Text.IsEmptyAndNull(bField) Then
			Dim tmpDic, tmpKey
			Select Case VarType(bValue)
				Case 0, 1: '// vbEmpty,vbNull
					Set tmpDic = Dicary(): PrDic.RemoveAll'// 清空所有配置参数
				Case 2, 3, 4, 5: '// vbInteger,vbLong,vbSingle,vbDouble
					Set tmpDic = System.Text.ToHashTable(Array("LIMIT:"&bValue))
				'Case 6: '// vbCurrency
				'Case 7: '// vbDate
				Case 8: '// vbString
					'// 如果目标参数的值为字符串时，将其转换为数组
					'// 其中对字符串用“|”符号进行分隔
					Dim tmpObj: Set tmpObj = System.Array.New
					tmpObj.Symbol = "|"
					tmpObj.Data = bValue
					Set tmpDic = System.Text.ToHashTable(tmpObj.ToArray)
					Set tmpObj = Nothing
				Case 9: '// vbObject
					Set tmpDic = bValue
				'Case 10: '// vbError
				'Case 11: '// vbBoolean
				Case 8192, 8194, 8204, 8209: '// 8192(Array),8204(vbVariant()),8209(Byte)
					Set tmpDic = System.Text.ToHashTable(bValue)
			End Select
			For Each tmpKey In tmpDic: PrDic(tmpKey) = tmpDic.Item(tmpKey): Next
			Set tmpDic = Nothing
		Else
			Select Case VarType(bValue)
				Case 0, 1:
					PrDic(bField) = ""
				'Case 2, 3, 4, 5, 6, 7, 8, 11:
				''	PrDic(bField) = bValue
				'Case 9:
				Case 8192, 8194, 8204, 8209:
					If UCase(bField) = "WHERE" Then
						PrDic(bField) = JoinWhere(bValue)
					ElseIf UCase(bField) = "FIELD" Then
						PrDic(bField) = System.Array.NewArray(bValue).J(",")
					Else PrDic(bField) = bValue(0) End If
				Case Else
					PrDic(bField) = bValue
			End Select
		End If
	End Property
	
	'// 获取参数集合
	'// 如果参数为空时，则返回一个DIC对象，否则返回目标项的值
	Public Property Get Parameters(ByVal bField)
		If System.Text.IsEmptyAndNull(bField) Then Set Parameters = PrDic _
		Else Parameters = PrDic(bField)
	End Property

	'// 拼接条件语句
	Private Function JoinWhere(ByVal bValue)
		Dim tmpObj: Set tmpObj = System.Array.NewHash(bValue)
		'// 判断是否存在逻辑判断符，如果存在则用此符号进行组装，否则用AND进行组装
		If LCase(tmpObj.HasIndex("_logic")) Then
			'// 获取逻辑判断符的值
			Dim blLogic: blLogic = tmpObj("_logic")
			'// 删除逻辑判断符 记录项
			tmpObj.Delete("_logic")
			JoinWhere = tmpObj.J(") "& blLogic &" (")
		Else JoinWhere = tmpObj.J(") AND (") End If
		Set tmpObj = Nothing
	End Function

	'// 新增数据
	Public Function Add(ByVal bValue)
		With System.Data
			PrDic("SQL") = .ToSQL(Array(PrTable, PrDic("FIELD"), PrDic("LIMIT")), PrDic("WHERE"), PrDic("ORDER"))
			Add = .Create(PrDic("SQL"), bValue)
		End With
	End Function

	'// 保存数据
	Public Function Save(ByVal bValue)
		With System.Data
			PrDic("SQL") = .ToSQL(Array(PrTable, PrDic("FIELD"), PrDic("LIMIT")), PrDic("WHERE"), PrDic("ORDER"))
			Save = .Update(PrDic("SQL"), bValue)
		End With
	End Function

	'// 查询数据
	Public Function [Select]()
		With System.Data
			PrDic("SQL") = .ToSQL(Array(PrTable, PrDic("FIELD"), PrDic("LIMIT")), PrDic("WHERE"), PrDic("ORDER"))
			Set [Select] = .Read(PrDic("SQL"))
		End With
	End Function

	'// 查询数据并分页
	Public Function Pager()
		With System.Data
			PrDic("SQL") = .ToSQL(Array(PrTable, PrDic("FIELD"), PrDic("LIMIT")), PrDic("WHERE"), PrDic("ORDER"))
			'// 将所有参数传递给分页类
			.Page.Parameters("") = Me.Parameters("")
			'// 对得到的结果进行行列对换
			Dim blList: blList = System.Array.Swap(.Page.Run)
			'// 返回数组，顺序依次为 [0]记录集列表，[1]分页导航码，[2]分页参数
			Pager = Array(blList, .Page.Out, .Page.Parameters(""))
		End With
	End Function

	'// 删除数据
	Public Function Delete(ByVal bValue)
		If Not System.Text.IsEmptyAndNull(bValue) Then PrDic("WHERE") = bValue
		PrDic("SQL") = System.Data.ToSQL(Array(PrTable, PrDic("FIELD"), PrDic("LIMIT")), PrDic("WHERE"), PrDic("ORDER"))
		Delete = System.Data.Delete(PrDic("SQL"))
	End Function

	'// 统计查询
	'// 统计数量，参数是统计的字段名（可选）
	Public Function Count(ByVal bValue)
		With System.Text
			Dim blField: blField = .IIF(Not .IsEmptyAndNull(bValue), bValue, "*")
			Dim blSQL: blSQL = "Select Count("&blField&") From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
		End With
		Count = System.Data.Read(PrDic("SQL"))(0)
	End Function
	'// 获取最大值，参数是要统计的字段名（必须）
	Public Function Max(ByVal bValue)
		With System.Text
			Dim blSQL: blSQL = "Select Max("&bValue&") From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
		End With
		Max = System.Data.Read(PrDic("SQL"))(0)
	End Function
	'// 获取最小值，参数是要统计的字段名（必须）
	Public Function Min(ByVal bValue)
		With System.Text
			Dim blSQL: blSQL = "Select Min("&bValue&") From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
		End With
		Min = System.Data.Read(PrDic("SQL"))(0)
	End Function
	'// 获取平均值，参数是要统计的字段名（必须）
	Public Function Avg(ByVal bValue)
		With System.Text
			Dim blSQL: blSQL = "Select Avg("&bValue&") From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
		End With
		Avg = System.Data.Read(PrDic("SQL"))(0)
	End Function
	'// 获取总分，参数是要统计的字段名（必须）
	Public Function Sum(ByVal bValue)
		With System.Text
			Dim blSQL: blSQL = "Select Sum("&bValue&") From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
		End With
		Sum = System.Data.Read(PrDic("SQL"))(0)
	End Function

	'// 字段值增长，只对单条记录进行更改
	'// bValue[:step]
	Public Function setInc(ByVal bValue)
		With System.Text
			Dim blField: blField = .Separate(bValue)
			Dim blStep: blStep = .IIF(Not .IsEmptyAndNull(blField(1)), blField(1), 1)
			Dim blSQL: blSQL = "Select Top 1 "&blField(0)&" From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
			Dim blSourceValue: blSourceValue = System.Data.Read(PrDic("SQL"))(0)
			blSourceValue = .IIF(IsNumeric(blSourceValue), blSourceValue, 0)
		End With
		setInc = System.Data.Update(PrDic("SQL"), Array(Array(blField(0), blSourceValue + blStep)))
	End Function

	'// 字段值减少，只对单条记录进行更改
	'// bValue[:step]
	Public Function setDec(ByVal bValue)
		With System.Text
			Dim blField: blField = .Separate(bValue)
			Dim blStep: blStep = .IIF(Not .IsEmptyAndNull(blField(1)), blField(1), 1)
			Dim blSQL: blSQL = "Select Top 1 "&blField(0)&" From "&PrTable&""
			PrDic("SQL") = .IIF(Not .IsEmptyAndNull(PrDic("WHERE")), (blSQL & " Where (" & PrDic("WHERE")) & ")", blSQL)
			Dim blSourceValue: blSourceValue = System.Data.Read(PrDic("SQL"))(0)
			blSourceValue = .IIF(IsNumeric(blSourceValue), blSourceValue, 0)
		End With
		setDec = System.Data.Update(PrDic("SQL"), Array(Array(blField(0), blSourceValue - blStep)))
	End Function
End Class
%>