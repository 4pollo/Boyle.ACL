<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统模板操作类]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

'// ---------------------------------------------------------------------------
'// 作者：Taihom(taihom@163.com)
'// 网址：http://www.cnblogs.com/taihom/
'// ---------------------------------------------------------------------------

Class Cls_Template
	Private dicLabel, tplXML, strSuffix
	Private strRootPath, strCharset, strTagHead, strRootXMLNode, strBlockDataAtr
	Private strTemplatePath, strTemplateFilePath, strTemplateHtml, strResultHtml
	Private strDateDiffTimeInterval, strTemplateCacheName, strTemplatePagePath
	Private strTemplateCachePath, intTemplateCacheType, intTemplateCacheTime
	Private intOpenAbsPath, strAppCacheName, strFileCachePath
	Private intLayout, strLayoutName, strLayoutItem, strLayoutFile

	'// 类初始化
	Private Sub Class_Initialize()

		'// 全局默认变量
		strCharset      = System.Charset 	'编码设置
		strSuffix       = ".html"			'设置模板文件的后缀名
		strTagHead      = "$"				'定义模板标签头
		strTemplatePath = "." 				'模板存放目录
		strRootXMLNode  = "//template"  	'模板根节点名称
		strBlockDataAtr = "name"        	'块赋值辅助的属性
		intOpenAbsPath  = 1					'输出结果是否使用绝对路径 (0不用,1用)
		
		strDateDiffTimeInterval = "s"       		'表示相隔时间的类型：d日 h时 n分钟 s秒
		intTemplateCacheType    = 0         		'缓存类型
		intTemplateCacheTime    = 10        		'缓存时间
		strTemplateCachePath    = "runtime/cache" 	'缓存目录
		
		intLayout     = 0 										'是否开启全局模板布局模式(0不开启,1开启)
		strLayoutName = "layout"								'全局模板布局的名称
		strLayoutItem = "{__CONTENT__}"							'全局模板布局的替换字符串
		strLayoutFile = TMPL_PATH & strLayoutName & strSuffix 	'全局模板布局的文件路径

		'Xml对象
		Set tplXML = xmlDom(Right(strRootXMLNode, Len(strRootXMLNode)-2))
		
		'使用到的字典对象
		Set dicLabel = Dicary()

		System.Error.E(300) = "模板文件不存在！"
	End Sub
	
	'// 类退出
	Private Sub Class_Terminate()
		Set tplXML   = Nothing
		Set dicLabel = Nothing
	End Sub
	
	'// 新建类实例
	Public Function [New]()
		Set [New] = New Cls_Template
	End Function
	
	'// 设置站点根目录路径
	Public Property Let setRootPath(ByVal strVal)
		strRootPath = strVal
	End Property
	
	'// 设置使用字符编码
	Public Property Let setCharset(ByVal strVal)
		strCharset = strVal
	End Property
	
	'// 设置单标签头
	Public Property Let setTagHead(ByVal strVal)
		strTagHead = LCase(Trim(strVal))
	End Property
	
	'// 输出结果时，将路径转换为绝对路径
	Public Property Let IsAbsPath(ByVal strVal)
		intOpenAbsPath = strVal
	End Property

	'// 设置模板存放路径
	Public Property Let setTemplatePath(ByVal strVal)
		strTemplatePath = strVal
	End Property
	Public Property Let Root(ByVal strVal)
		strTemplatePath = strVal
	End Property
	Public Property Get Root()
		Root = strTemplatePath
	End Property
	
	'// 设置模板文件路径
	Public Property Let setTemplateFile(ByVal strVal)
		strTemplatePagePath = strVal
		strTemplateFilePath = System.IO.FormatFilePath(strRootPath & strTemplatePath & strTemplatePagePath)
		'文件缓存路径
		strFileCachePath = strTemplateCacheName & "/" & Replace(strTemplatePagePath, "."&System.IO.FileExts(strTemplatePagePath), "")
		'内存缓存的名称
		strAppCacheName = strTemplateCacheName & "_" & intTemplateCacheTime & "_" & intTemplateCacheType & "_" & strTemplatePagePath
		'设置路径后立即加载模板
		call loadCacheTemplate()
	End Property

	'// 同时对模板路径和文件进行设置
	Public Property Let File(ByVal strVal)
		setTemplateFile = strVal & strSuffix
	End Property
	Public Property Get File()
		File = strTemplatePagePath
	End Property
	'// 定义模板文件的后缀名
	Public Property Let Suffix(ByVal strVal)
		If Left(strVal, 1) <> "." Then strSuffix = "."&strVal _
		Else strSuffix = strVal
	End Property
	Public Property Get Suffix()
		Suffix = strSuffix
	End Property
	'// 定义是否开启全局模板布局
	Public Property Let Layout(ByVal strVal)
		intLayout = System.Text.ToBoolean(strVal)
	End Property
	Public Property Get Layout()
		Layout = intLayout
	End Property
	'// 定义是否全局模板布局的模板名称
	Public Property Let LayoutName(ByVal strVal)
		strLayoutName = strVal
	End Property
	Public Property Get LayoutName()
		LayoutName = strLayoutName
	End Property
	'// 定义是否全局模板布局的替换字符串
	Public Property Let LayoutItem(ByVal strVal)
		strLayoutItem = strVal
		'// 设置全局模板布局的文件路径，这里用到了一个全局变量，即模板根目录路径
		strLayoutFile = TMPL_PATH & strLayoutItem & strSuffix
	End Property
	Public Property Get LayoutItem()
		LayoutItem = strLayoutItem
	End Property
	
	'参数1: 缓存的名字,每个页面不能相同
	'参数2: 0=都不缓存,1=内存缓存,2=文件缓存(缓存会缓存数据跟模板,开启缓存必须要有一个缓存名字)
	'参数3: 缓存时间，单位是默认是秒
	Public Property Let setCache(ByVal strVal)
		Dim arr: arr = expSplit(strVal, "\s*,\s*")
		Select Case UBound(arr, 1)
		Case 0
			strTemplateCacheName = arr(0)
		Case 1
			strTemplateCacheName = arr(0)
			intTemplateCacheType = CInt(arr(1))
		Case 2
			strTemplateCacheName = arr(0)
			intTemplateCacheType = CInt(arr(1))
			intTemplateCacheTime = CInt(arr(2))
		Case 3
			strTemplateCacheName = arr(0)
			intTemplateCacheType = CInt(arr(1))
			intTemplateCacheTime = CInt(arr(2))
			strTemplateCachePath = arr(3)
		End Select
		If intTemplateCacheTime <= 0 Then
			intTemplateCacheType = 0
		End If
		'//设置文件缓存的保存路径，这里用到全局变量APP_PATH，即项目路径
		System.Cache.SavePath = APP_PATH&strTemplateCachePath
	End Property
	
	'赋值
	Public Property Let d(ByVal strTag, ByVal strVal)
		Dim i, ary: ary = expSplit(strTag, "\s*,\s*")
		For i = 0 To Ubound(ary)'多标签赋值
			strTag = LCase(ary(i))
			If strTag = strTagHead Then
				'Dim tmpDic: Set tmpDic = Dicary()
				Select Case TypeName(strVal)
					Case "Recordset"'记录集
						If strVal.State And Not strVal.Eof Then
							Set dicLabel = RsToDic(strVal, dicLabel)
						End If
					Case "Dictionary"'// 如果集合中已经存在标签，则进行追加。
						Set dicLabel = RsToDic(strVal, dicLabel)
					Case "Variant()"'如果传递的是数组
						If Ubound(strVal) = 1 Then
							Select Case TypeName(strVal(0))
								Case "Recordset"
									If strVal(0).State And Not strVal(0).eof Then
										Set dicLabel = RsToDic(strVal(0), dicLabel)
									End If
								Case "Variant()"
									Set dicLabel = RsToDic(strVal(0), dicLabel)
							End Select
							Dim aryField: aryField = expSplit(strVal(1), "\s*,\s*")'字段序列
							If TypeName(aryField)="Variant()" Then Set dicLabel = RedimField(dicLabel, aryField)'重命名字段
						End If
				End Select
				'// [BOYLE.ACL]如果集合中已经存在标签，则进行追加。
				'If Not System.Text.IsEmptyAndNull(dicLabel) Then
				''	Dim tmpKey: For Each tmpKey In tmpDic: dicLabel(tmpKey) = tmpDic.Item(tmpKey): Next
				'Else Set dicLabel = tmpDic End If
				'Set dicLabel = tmpDic
				'Set tmpDic = Nothing
			Else'普通赋值,支持字典，普通数据(字段值、字符串、数字等)
				Select Case TypeName(strVal)
					Case "Dictionary", "Recordset"
						Set dicLabel(strTag) = strVal
					Case Else
						dicLabel(strTag) = strVal
				End Select
			End If
		Next
	End Property

	'生成静态页面(路径,页面名称)
	Public Property Let Create(ByVal param)
		Dim strFilePath, strContents
		If TypeName(param) = "Variant()" Then'传递数组
			Select Case Ubound(param)
				Case 0'Array(createpath+pagename)
					strFilePath = param(0)
					strContents = getHtml
				Case 1'Array(createpath+pagename,content)
					strFilePath = param(0)
					strContents = param(1)
				Case Else'Array(createpath+pagename,content,charset)
					strFilePath = param(0)
					strContents = param(1)
					strCharset  = param(2)
			End Select
		Else '文件路径+文件名
			strFilePath = param
			strContents = getHtml
		End If
		System.IO.Charset = strCharset
		System.IO.Save strFilePath, strContents
	End Property
	
	'设置节点属性
	Public Property Let setAttr(ByVal strPath,ByVal v)
		Attr(strPath) = v
	End Property
	Public Property Let Attr(ByVal strPath,ByVal v)
		SetLabelAttr LCase(strPath),v
	End Property
	
	'获取节点属性
	Public Property Get getAttr(ByVal strPath)
		Attr strPath
	End Property
	Public Property Get Attr(ByVal strPath)
		Dim i,ary,node
		ary = selectLabelNode(LCase(strPath))'选择标签节点
		If IsArray(ary) = False Then Exit Property
		
		Select Case LCase(ary(3))
		Case ":body"
			Set node = tplXML.selectNodes(ary(4) & "/body")
		Case ":empty",":null",":eof"
			Set node = tplXML.selectNodes(ary(4) & "/null")
		Case ":html"
			Set node = tplXML.selectNodes(ary(4) & "/html")
		Case Else
			If Len(ary(2)) Then
				Set node = tplXML.selectNodes(ary(4)&"/@"&ary(2))
			Else'如果没有属性路径就返回节点的所有属性
				Set node = tplXML.selectNodes(ary(4))
				Redim tagAttr(node.Length)
				For i = 0 to node.Length - 1
					Set tagAttr(i) = getBlockAttr(node(i))
				Next
				Attr = tagAttr
				Exit Property
			End If
		End Select
		
		If IsObject(node) Then
			If node.Length Then 
				Redim tagAttr(node.Length)
				For i = 0 to node.Length - 1
					tagAttr(i) = node(i).nodeTypedValue
				Next

				'如果只有一个结果，就返回这个结果
				If Ubound(tagAttr) = 1 Then
					Attr = tagAttr(0)
				Else'如果有多个结果就返回数组
					Attr = tagAttr
				End If
			End If
			Set node = Nothing
		End If
	End Property
	
	'// 获得标签所有的值
	Public Property Get GetLabelValues(ByVal strVal)
		If IsObject(GetLabVal(strVal)) Then Set GetLabelValues = GetLabVal(strVal) _
		Else GetLabelValues = GetLabVal(strVal)
	End Property
	Public Property Get GetLabVal(ByVal strVal)
		If LCase(strVal) = LCase(strTagHead) Then'如果返回所有值对象
			Set GetLabVal = dicLabel
		Else
			If IsObject(dicLabel(strVal)) Then Set GetLabVal = dicLabel(strVal) _
			Else GetLabVal = dicLabel(strVal)
		End If
	End Property
	
	'// 输出部分
	Public Property Get GetHtml
		Select Case intTemplateCacheType
			Case 3'结果内存缓存
				Dim CacheName: CacheName = strAppCacheName & "getHtml"
				System.Cache.Item(CacheName).Expires = intTemplateCacheTime/60
				If System.Cache.Item(CacheName).Ready Then
					strResultHtml = System.Cache.Item(CacheName)
				Else
					Call AnalysisTemplate()
					System.Cache.Item(CacheName) = strResultHtml
					System.Cache.Item(CacheName).SaveApp
				End If
			Case 4'结果文件缓存
				System.Cache.Item(strFileCachePath).Expires = intTemplateCacheTime/60
				If System.Cache.Item(strFileCachePath).Ready Then
					strResultHtml = System.Cache.Item(strFileCachePath)
				Else
					Call AnalysisTemplate()
					System.Cache.Item(strFileCachePath) = strResultHtml
					System.Cache.Item(strFileCachePath).Save
				End If
			Case Else
				Call AnalysisTemplate()
		End Select
		
		'返回执行时间和数据库执行次数
		strResultHtml = System.Text.ReplaceX(strResultHtml, "\{runtime\s*\/?\}|(\<\!--runtime--\>)(.*?)\1", "<"&"!--runtime-->"&System.End&"<"&"!--runtime-->" )
		strResultHtml = System.Text.ReplaceX(strResultHtml, "\{queries\s*\/?\}|(\<\!--queries--\>)(.*?)\1", "<"&"!--queries-->"&System.Queries&"<"&"!--queries-->" )
		getHtml = strResultHtml
	End Property
	
	'//输出模板部分
	Public Property Get Display
		Response.Write(getHtml)
	End Property

	'// ----------------------私有函数部分----------------------	
	'xmlDom对象
	Private Function XmlDom(ByVal root)
		Set XmlDom = Server.CreateObject("MSXML2.DOMDocument")
		'.Async选项设置成'False'，是为了告诉浏览器中的XML解析器：一边读取XML文档，一边进行数据显示
		XmlDom.Async = False
		If Len(root) > 0 Then
			'创建一个节点对象
			XmlDom.appendChild(XmlDom.CreateElement(root))
			'添加xml头部
			Dim head: Set head = XmlDom.CreateProcessingInstruction("xml","version=""1.0"" encoding="""&strCharset&"""")
			XmlDom.insertBefore head, XmlDom.childNodes(0)
		End If
	End Function
	
	'转义正则字符
	Private Function expEncode(ByVal sText)
		Dim i, ary: ary = Split(". * + ? | ( ) { } ^ $ :", " ")
		sText = Replace(sText, "\" , "\\")
		For i = 0 to Ubound(ary)
			sText = Replace(sText, ary(i) , "\"&ary(i))
		Next
		expEncode = sText
	End Function

	'ASP的正则expSplit
	Private Function expSplit(ByVal a, ByVal b)
		Dim Match, SplitStr : SplitStr = a
		Dim Sp : Sp = "#taihom.com@"
			For Each Match in System.Text.MatchX(a, b)
				SplitStr = Replace(SplitStr, Match.Value, Sp, 1, -1, 0)
			Next
		expSplit = Split(SplitStr, Sp)
	End Function
	
	'功能:返回指定数组的维数
	Private Function GetArrayDimension(ByVal aryVal)
		On Error Resume Next
		GetArrayDimension = -1
		If Not IsArray(aryVal) Then
			Exit Function
		Else
			Dim i, iDo
			For i = 1 To 4
				iDo = UBound(aryVal, i)
				If Err Then Err.Clear: Exit Function _
				Else GetArrayDimension = i
			Next
		End If
	End Function
	
	'加载或者缓存模板
	Private Sub LoadCacheTemplate
		'缓存类型 0=不缓存,1=内存缓存,2=文件缓存
		Select Case intTemplateCacheType
		Case 0'不缓存
			Call Load()
		Case 1,3'1=模板内存缓存,3=结果内存缓存
			Dim TplHtmlCache, TplXmlCache
			TplHtmlCache = strAppCacheName & ".TPLHTML.APP"
			TplXmlCache = strAppCacheName & ".TPLXML.APP"
			System.Cache.Item(TplHtmlCache).Expires = intTemplateCacheTime/60
			System.Cache.Item(TplXmlCache).Expires = intTemplateCacheTime/60
			If System.Cache.Item(TplHtmlCache).Ready Then
				strTemplateHtml = System.Cache.Item(TplHtmlCache)
				Set tplXML = XmlDom("")
				tplXML.loadXML(System.Cache.Item(TplXmlCache))
			Else
				Call Load()
				System.Cache.Item(TplHtmlCaChe) = strTemplateHtml
				System.Cache.Item(TplXmlCache) = TplXml.xml
				System.Cache.SaveAppAll
			End If
		Case 2,4'2=模板文件缓存,4=结果文件缓存
			Dim CacheName : CacheName = strFileCachePath & ".xml"
			System.Cache.Item(CacheName).Expires = intTemplateCacheTime/60
			If System.Cache.Item(CacheName).Ready Then
				Set tplXML = XmlDom("")
				tplXML.loadXML(System.Cache.Item(CacheName))
				strTemplateHtml = tplXML.SelectSingleNode(strRootXMLNode).LastChild.data
			Else
				Call Load()
				System.Cache.Item(CacheName) = tplXML.xml
				System.Cache.Item(CacheName).Save
			End If
		'Case 3,4'3=结果内存缓存,'4=结果文件缓存
		'	Call GetHtml()
		End Select
	End Sub
	
	Private Sub Load()'读取模板文件
		If Not System.IO.ExistsFile(strTemplateFilePath) Then System.Error.Raise 300: Exit Sub

		'// 判断是否开启全局布局模式，如果开启则读取文件
		Dim blLayoutHtml: blLayoutHtml = strLayoutItem
		If intLayout Then blLayoutHtml = System.IO.Read(System.IO.FormatFilePath(strLayoutFile))
		'// 将全局模板文件并入模板文件中，进行解析
		strTemplateHtml = Replace(blLayoutHtml, strLayoutItem, System.IO.Read(strTemplateFilePath), 1, -1, 0)
		strTemplateHtml = LoadInclude(strTemplateHtml, strTemplateFilePath)
		strTemplateHtml = System.Text.ReplaceX(System.Text.ReplaceX(strTemplateHtml, "\<\!\-\-\s*\{","{"),"\}\s*\-\-\>", "}")

		'// 使用绝对路径
		strTemplateHtml = System.Text.IIF(CBool(intOpenAbsPath), AbsPath(strTemplateHtml), strTemplateHtml)

		'编译模板，并且用XML存储模板标签节点
		Dim XmlRoot: Set XmlRoot = tplXML.SelectSingleNode(strRootXMLNode)
		CompileTemplate Array(strTemplateHtml, strTagHead), XmlRoot
		'保存模板到XML
		XmlRoot.appendChild(tplXML.CreateCDATASection(strTemplateHtml))
		strTemplateHtml = XmlRoot.LastChild.Data
	End Sub
	
	'模板的include支持
	Private Function LoadInclude(ByVal strHtml,ByVal strPath)
		Dim Match, incPath, html: html = strHtml
		For Each Match In System.Text.MatchX(strHtml, "{include\s*([\('""])?\s*(.*?)\1}")
			incPath = System.IO.FormatFilePath(strRootPath & strTemplatePath & Match.SubMatches(1))
			If strPath <> incPath Then html = Replace(html, Match.Value, LoadInclude(System.IO.Read(incPath), incPath), 1, -1, 0) _
			Else html = Replace(html, Match.Value, "", 1, -1, 0)
		Next
		LoadInclude = html
	End Function
	
	'编译模板
	'参数：模板内容,标签头,XML节点路径
	Private Sub CompileTemplate(ByVal aryVal,ByVal nodeDOM)
		If Len(aryVal(1))=0 Then Exit Sub End If
		Dim Match,Matches,strPattern
		Dim arrayTags(10) '定义一个数组，把模板的标签参数保存调用
		strPattern = "\{("&ExpEncode(LCase(aryVal(1)))&")([a-zA-Z0-9:_]+)?\s*?([\s\S]*?)\/?\}[\n|\s|\t]*?(?:[\n]*?([\s\S]*?)[\n|\s|\t]*?(\{/\1\2\}))?"
		'解析标签
		For Each Match in System.Text.MatchX(aryVal(0), strPattern)
			arrayTags(0) = LCase(Match.SubMatches(0)) ' 标签头
			arrayTags(1) = LCase(Match.SubMatches(1)) ' 标签名称
			arrayTags(2) = Match.SubMatches(2) ' 标签属性
			arrayTags(3) = Match.SubMatches(3) ' 闭合部分的内容
			arrayTags(4) = ""                  ' empty标签
			arrayTags(5) = arrayTags(3)        ' 仅循环体部分,不包含empty
			arrayTags(6) = System.Text.IIF(Len(Match.SubMatches(4))+Len(Match.SubMatches(3)),1,0) ' 如果是闭合标签，并且有模板内容，闭合标签才有效
			arrayTags(7) = Match.Value '模板内容

			'如果是有结束标签,表示这个是一个闭合标签
			If arrayTags(6) Then
			Dim closeTags : closeTags = GetCloseBlock(Array(arrayTags(3),arrayTags(1)))
				arrayTags(4) = closeTags(0)    ' empty标签
				arrayTags(5) = closeTags(1)    ' 仅循环体部分,不包含empty
				arrayTags(8) = System.Text.ReplaceX( getBlockAttr(nodeDOM)("nodepath") & "/" & arrayTags(1),"^\/","")'节点路径
			End If
			'创建节点
			nodeDOM.appendChild(GetTemplateNode(arrayTags))
		Next
	End Sub
	
	'解析模板
	'模板输出的思路是，遍历模板标签节点，根据编译的节点信息来输出值
	Private Sub AnalysisTemplate
		Dim node: Set node = tplXML.selectNodes(strRootXMLNode)'从根目录开始遍历模板标签节点
		
		strResultHtml = AnalysisBlockLabel(node(0).lastchild.nodeTypedValue, node(0).childNodes, strTagHead, dicLabel)'循环以及嵌套循环标签		
		strResultHtml = ReturnLabelValues(strResultHtml, strTagHead, dicLabel, 1)'单标签
		strResultHtml = ExecuteTemplate(ReturnIfLabel(strResultHtml, strTagHead, dicLabel))
		Set node = Nothing
	End Sub
	
	'解析标签，获取值
	'参数：代码、节点、标签前缀、字典数据(用来支持标签值的调用)
	Private Function AnalysisBlockLabel(ByVal strHtml,ByVal node,ByVal strHead,ByVal objDIC)
		If Len(strHtml) = 0 Then Exit Function End If
		'由于是从根节点开始遍历，所以不用考虑多个标签相同的情况，所以只要遍历根结点的子节点就可以了
		Dim html : html = strHtml
		Dim i
		For i = 0 To node.Length - 1
		'遍历所有子节点,遇到循环就递归调用
			If node(i).childNodes.Length > 3 Then
			
				Dim DicData : Set DicData = Dicary()
				Dim aryLabel: aryLabel = GetLabelNode(node(i))'提取节点值
				
				'获取节点路径
				If Len(aryLabel(5)(strBlockDataAtr)) Then
					DicData("path") = aryLabel(5)("nodepath") & "[" & strBlockDataAtr & "=" & aryLabel(5)(strBlockDataAtr) & "]"
				Else
					DicData("path") = aryLabel(5)("nodepath")
				End If
				'根据节点路径获得值
				If dicLabel.Exists(DicData("path")) Then
					If IsObject(dicLabel(DicData("path"))) Then
						Set DicData("data") = dicLabel(DicData("path"))
					Else
						DicData("data") = dicLabel(DicData("path"))
					End If
				Else'如果没有赋值,就根据节点设置的获得值
					Dim sql,conn,rs
					sql = aryLabel(5)("sql")
					conn= aryLabel(5)("conn")
						
					sql  = System.Text.ReplaceX(sql,"^(\w+)\((?:\w+)\s*,\s*(?:\w+)\)$","$1(aryLabel(0),aryLabel(5))")'动态SQL赋值
					If Right(sql,25) = "(aryLabel(0),aryLabel(5))" Then
						sql = Eval(sql)
					End If
						
					sql  = ReturnLabelValues(sql,strHead,objDIC,0)'获取Sql
					conn = ReturnLabelValues(conn,strHead,objDIC,0)'获取CONN

					If Len(sql) > 6 Then'如果SQL不为空
						'On Error Resume Next
						If Len(conn) > 2 Then Set conn = Eval(conn) Else Set conn = System.Data.Connection
						Set DicData("data") = System.Data.Read(sql)
					End If
				End If
				
				DicData("tagHead") = aryLabel(0)&"."
				
				'获得处理值,返回的数据只有两种情况,一种格式数组，一种是其他格式
				Set DicData = GetBlockData(DicData)
					'DicData("field") '如果需要重定义字段名
					'DicData("dr") '设置有dr函数
					'DicData("eof") = 1 表示没有数据
					'DicData("returndata") 返回的数据
					'DicData("data") 原来的数据
				
				'数据准备结束
				If Not DicData.Exists("dr") Then'如果没有设置渲染块
					DicData("dr") = System.Text.ReplaceX(aryLabel(5)("dr"),"\s*([a-zA-Z0-9]+)\(([a-zA-Z0-9]+)\)\s*","$1(dicRS)")'数据渲染
				End If

				'数据处理
				Dim returnData : returnData = DicData("returndata")
				Dim returnHtml : returnHtml = ""
				Dim dicRS , k : k = 0
				If DicData("eof") Then'如果没有数据
					returnHtml = ReturnLabelValues(aryLabel(4),strHead,objDIC,1)
				Else'如果有数据
					Select Case TypeName(returnData)
						Case "Variant()"'统一返回数组
							For k = 0 To Ubound(returnData)
								'获取字段值
								Set dicRS = returnData(k) : dicRS("i") = k + 1
								'重命名字段
								If TypeName(DicData("field"))="Variant()" Then Set dicRS = RedimField(dicRS, DicData("field"))
								'数据重定义或渲染
								If Right(DicData("dr"),7)="(dicRS)" Then Eval(DicData("dr"))
								'返回块的值
								returnHtml = returnHtml &_
								AnalysisBlockLabel(ReturnLabelValues(aryLabel(3),DicData("tagHead"),dicRS,1),node(i).childNodes,DicData("tagHead"),dicRS)'递归循环
								returnHtml = ReturnIfLabel(returnHtml,DicData("tagHead"),dicRS)'搞定IF比较值
							Next
							'returnHtml = DicData("blockHtml")
						Case Else'其他类型
							returnHtml = returnData
					End Select
				End If

				'标签替换
				html = Replace(html,aryLabel(2),returnHtml,1,-1,0)
				Set DicData = Nothing
			End If
		Next
		AnalysisBlockLabel = html
	End Function
	
	'获得块的值以及类型
	'传递数据，返回一维数据，元素一定要是字典数据否则不能处理
	Private Function GetBlockData(ByVal DicData)
		'检测块值的类型
		Dim aryTemp,aryData,dic,rs,returnData
		Dim recIndex,fldIndex
		Select Case TypeName(DicData("data"))
			Case "Recordset"'如果块传值是记录集
				Set rs = DicData("data")
				If rs.Eof Then
					DicData("eof") = 1
				Else
					aryTemp = rs.getRows()
					ReDim aryData(Ubound(aryTemp, 2), Ubound(aryTemp, 1))
					ReDim returnData(Ubound(aryTemp, 2))					
					For recIndex = 0 To UBound(aryTemp,2)'行
						Set dic = Dicary()
						For fldIndex = 0 To UBound(aryTemp)'字段
							'aryData(recIndex,fldIndex) = aryTemp(fldIndex,recIndex)&""
							dic(LCase(rs.Fields(fldIndex).Name)) = aryTemp(fldIndex, recIndex)&""
							dic(fldIndex) = aryTemp(fldIndex, recIndex)&""
						Next
						Set returnData(recIndex) = dic'返回数组
					Next
				End If
			Case "Variant()"'如果传递的是数组
				If Ubound(DicData("data")) = 1 Then
					DicData("field") = ExpSplit(DicData("data")(1),"\s*,\s*")'字段序列
					Select Case TypeName(DicData("data")(0))
					Case "Recordset"'数据集
						Set rs = DicData("data")(0)
						If rs.Eof Then
							DicData("eof") = 1
						Else
							aryTemp = rs.getRows
							ReDim aryData(Ubound(aryTemp, 2), Ubound(aryTemp))
							ReDim returnData(Ubound(aryTemp, 2))
							For recIndex = 0 To UBound(aryTemp, 2)'行
								Set dic = Dicary()
								For fldIndex = 0 To UBound(aryTemp)'字段
									'aryData(recIndex,fldIndex) = aryTemp(fldIndex,recIndex)&""
									dic(LCase(rs.Fields(fldIndex).Name)) = aryTemp(fldIndex, recIndex) & ""
									dic(fldIndex) = aryTemp(fldIndex, recIndex) & ""
								Next
								Set returnData(recIndex) = dic'返回数组
							Next
						End If
					Case "Variant()", "Cls_Array"	'数组,超级数组类
						If TypeName(DicData("data")(0)) = "Cls_Array" Then
							aryTemp  = DicData("data")(0).Data
						Else
							aryTemp  = DicData("data")(0)
						End If
						Dim arycount : arycount = GetArrayDimension(aryTemp)						
						If arycount = 1 Then'如果是一维数组
							If Ubound(aryTemp) = 0 Then
								DicData("eof") = 1
							Else
								ReDim returnData(0)
								Set dic = Dicary()
								For fldIndex = 0 To UBound(aryTemp) '字段
									dic(fldIndex) = aryTemp(fldIndex) & ""
								Next
								Set returnData(0) = dic
							End If
						ElseIf arycount = 2 Then'二维数组
							If Ubound(aryTemp, 1)=0 Then
								DicData("eof") = 1
							Else
								ReDim returnData(Ubound(aryTemp,1))
								For recIndex = 0 To Ubound(aryTemp,1)
									Set dic = Dicary()
									For fldIndex = 0 To Ubound(aryTemp,2)'二级循环数据赋值
										dic(fldIndex) = aryTemp(recIndex, fldIndex) & ""'字段下标
									Next
									Set returnData(recIndex) = dic'返回数组
								Next
							End If
						Else returnData = Null End If
					Case Else
						returnData = aryTemp
					End Select
				End If
			Case "Dictionary"'如果传递的是字典
				If DicData("data").Count = 0 Then
					DicData("eof") = 1
				Else
					ReDim returnData(0)
					Set returnData(0) = DicData("data")
				End If
			Case Else'其他数据类型，主要是 字符、数字等可以直接输出的类型
				returnData = DicData("data")
		End Select
		'设置返回值函数
		DicData("returndata") = returnData
		
		Set GetBlockData = DicData
		Set dic = Nothing
	End Function

	'// 格式化值输出,参数：值，属性
	Private Function FormatValues(ByVal strVal,ByVal dicAttr)
		Dim return : return = strVal
		Dim key,val
		For Each key In dicAttr''遍历节点属性节点,根据节点的属性返回值
			val = ReturnLabelValues(dicAttr(key), strTagHead, dicLabel, 0)
			Select Case (LCase(key))
			Case "dateformat":'日期格式化
				return = System.Text.FormatTime(return, val)
			Case "len","length"
				return = System.Text.IIF(Len(val),Left(return,val),return)
			Case "cut"
				return = System.Text.Cut(return, val)
			Case "return"
				Dim str,i : val = Split(LCase(val),",")
				For i=0 To Ubound(val)
					Select Case val(i)
					Case "urlencode":
						return = Server.URLEncode(return)
					Case "htmlencode":
						return = System.Text.HTMLEncode(return)
					Case "htmldecode":
						return = System.Text.HTMLDecode(return)
					Case "clearhtml","removehtml":
						return = System.Text.RemoveHtml(return)
					Case "clearspace":
						return = System.Text.ReplaceX(return,"[\n\t\r|]|(\s+|&nbsp;|　)+", "")
					Case "clearformat":'清除所有格式
						return = System.Text.ReplaceX(return,"<[^>]*>|[\n\t\r|]|(\s+|&nbsp;|　)+", "")
					End Select
					str = str & return
				Next
				return = str
			End Select
		Next 
		FormatValues = return
	End Function
	
	'重定义字段数据
	Private Function RedimField(ByVal DicData, aryField)
		Dim i: For i = 0 To Ubound(aryField)
			If DicData.Exists(i) Then DicData(LCase(aryField(i))) = DicData(i)
		Next
		Set RedimField = DicData
	End Function
	
	'记录集转为字典
	Private Function RsToDic(ByVal data, ByVal dic)
		Dim i
		Select Case TypeName(data)
			Case "Recordset"'数据集
				For i = 0 To data.Fields.Count - 1 '字段序列
					dic(LCase(data.Fields(i).Name)) = data(i) & ""'字段名
					dic(i) = data(i) & ""'字段下标
				Next
			Case "Dictionary"'// 如果集合中已经存在标签，则进行追加。
				Dim tmpKey: For Each tmpKey In data: dic(tmpKey) = data.Item(tmpKey): Next
			Case "Variant()"'数组
				For i = 0 To Ubound(data) '字段序列
					dic(i) = data(i) & "" '字段下标
				Next
		End Select
		Set RsToDic = dic
	End Function

	'ExecuteTemplate
	Private Function ExecuteTemplate(ByVal strHtml)
		Dim html : html = strHtml
		Dim Matchs: Set Matchs = System.Text.MatchX(html, "\{(?:if)\s+([^}]*?)?\}")
		If Matchs.Count Then
			html = System.Text.ReplaceX(html, "\{(?:if)\s+([^}]*?)?\}", "<"&"%If $1 Then%"&">")
			html = System.Text.ReplaceX(html, "\{(?:elseif|ef)\s+([^}]*?)?\}", "<"&"%ElseIf $1 Then%"&">")
			html = System.Text.ReplaceX(html, "\{(?:else\s+if)\s+([^}]*?)?\}", "<"&"%Else If $1 Then%"&">")
			html = System.Text.ReplaceX(html, "\{else\s*\}", "<"&"%Else%"&">")
			html = System.Text.ReplaceX(html, "\{/if\}", "<"&"%End If%"&">")
		End If
		'Execute(html)
		Set Matchs = System.Text.MatchX(html, "\<"&"%([\s\S]*?)%"&"\>")
		If Matchs.Count Then'ASP代码支持，还不是那么完美,如果要解决，就要在下面的代码里面做处理
		Dim tmp : tmp = expSplit(html, "\<"&"%([\s\S]*?)%"&"\>")			
			Dim htm : htm = "Dim str : str = """"" & vbcrlf
			Dim i: For i = 0 To UBound(tmp)
				If Not System.Text.IsEmptyAndNull(System.Text.ReplaceX(tmp(i), "[\n\t\r|]|(\s+|&nbsp;|　)+", "")) Then
					tmp(i) = Replace(Replace(tmp(i), "<"&"%", "&lt;%"), "%"&">", "%&gt;")
					htm = htm & "str = str & tmp("&i&")" & vbcrlf
				End If
				If i <= (Matchs.Count - 1) Then htm = htm & Matchs(i).SubMatches(0) & vbcrlf
			Next			
			Execute(htm): html = str
		End If
		Set Matchs = Nothing
		ExecuteTemplate = html
	End Function
	
	'IF
	Private Function ReturnIfLabel(ByVal strHtml,ByVal strHead,ByVal dicRS)
		Dim Match, html: html = strHtml
		For Each Match in System.Text.MatchX(strHtml, "\{(?:if|elseif|ef)\s+([^}]*?)?\}")
			html = Replace(html, Match.Value, returnLabelValues(Match.Value, strHead, dicRS, 0))
		Next
		ReturnIfLabel = html
	End Function
	
	'// 标签属性值替换输出
	Private Function ReturnLabelValues(ByVal strVal, ByVal strHead, ByVal dicObj, ByVal key)
		Dim return, html : html = strVal
		Dim val, Match
		Dim Pattern(2)
		Pattern(0) = "\((?:" & expEncode(LCase(strHead)) &"){1}([a-zA-Z0-9\/_]+)((?:\[@?(?:\w+=.*?)?\])?\.?(?:\w+)?(?:\:\w+)?)?(\s+[^)][\s\S]*?)?\s*\)"'()标签
		Pattern(1) = "\{(?:" & expEncode(LCase(strHead)) &"){1}([a-zA-Z0-9\/_]+)((?:\[@?(?:\w+=.*?)?\])?\.?(?:\w+)?(?:\:\w+)?)?(\s+[^}][\s\S]*?)?\s*\}"'{}标签
		'(0)'标签名  (1)'路径  (2)'属性
		For Each Match in System.Text.MatchX(strVal, Pattern(key))
			If Len(Match.SubMatches(1)) Then'如果是通过路径获取属性
				return = GetAttr(Match.SubMatches(0)&Match.SubMatches(1))
			Else
				return = dicObj(LCase(Match.SubMatches(0)))
			End If
			If Len(Match.SubMatches(2)) > 1 Then
				return = FormatValues(return, GetBlockAttr(Match.SubMatches(2)))
			End If			
			html = Replace(html, Match.Value, return, 1, -1, 0)
		Next
		ReturnLabelValues = html
	End Function
	
	'// 返回一个标签节点的信息
	Private Function GetLabelNode(ByVal node)
		Dim aryLabel(6)
		aryLabel(0) = node.nodeName '节点名称
		If node.childNodes.Length < 3 Then
			aryLabel(1) = node.childNodes(0).nodeTypedValue '0=strAttr
			aryLabel(2) = node.childNodes(1).nodeTypedValue '1=strHtml
		End If
		If node.childNodes.Length > 3 Then
			aryLabel(1) = node.childNodes(0).nodeTypedValue '0=strAttr
			aryLabel(2) = node.childNodes(1).nodeTypedValue '1=strHtml
			aryLabel(3) = node.childNodes(2).nodeTypedValue '2=strBody
			aryLabel(4) = node.childNodes(3).nodeTypedValue '3=strEmpty
		End If
		Set aryLabel(5) = GetBlockAttr(node)'标签节点的所有属性
		GetLabelNode = aryLabel
	End Function
	
	'// 创建一个模板节点
	Private Function GetTemplateNode(ByVal arrayTags)
		'XML操作部分
		Dim subNode0,subNode1,subNode2,subNode3,subNode4,subNode5
		Set subNode0 = tplXML.CreateElement(LCase(Trim(arrayTags(1))))
		Set subNode1 = tplXML.CreateElement("attr") : subNode1.appendChild(tplXML.createCDATASection(arrayTags(2)))'标签属性
		Set subNode2 = tplXML.CreateElement("html") : subNode2.appendChild(tplXML.createCDATASection(arrayTags(7)))'模板内容
		Set subNode3 = tplXML.CreateElement("body") : subNode3.appendChild(tplXML.createCDATASection(arrayTags(5)))'循环体部分
		Set subNode4 = tplXML.CreateElement("null") : subNode4.appendChild(tplXML.createCDATASection(arrayTags(4)))'empty标签		
		'设置节点的属性
		Dim keys, tagAttr: Set tagAttr = GetBlockAttr(arrayTags(2))'提取属性部分，名=值
		'添加子节点
		subNode0.appendChild(subNode1)
		subNode0.appendChild(subNode2)		
		If arrayTags(6) Then'如果是闭合标签
			subNode0.appendChild(subNode3)
			subNode0.appendChild(subNode4)
			subNode0.SetAttribute "nodepath" , arrayTags(8) '辅助路径属性
			If Len(arrayTags(2)) > 1 Then
				Dim strSql: strSql = System.Text.ReplaceX(tagAttr("sql"), "^(\w+)\(\s*(\w+)\s*,\s*(\w+)\s*\)$", "$1(Tags,tagAttr)")'找SQL
				If Right(strSql, 14) = "(Tags,tagAttr)" Then strSql = Eval(strSql) End If
				subNode0.SetAttribute "sql" , strSql'SQL属性
			End If			
			'递归调用，这里是实现嵌套循环的关键
			compileTemplate Array(arrayTags(3),arrayTags(0)),subNode0 
		End If		
		'添加属性到节点中
		For Each keys in tagAttr
			subNode0.SetAttribute keys,tagAttr(keys)
		Next		
		Set getTemplateNode = subNode0
	End Function
	
	'// 分离EMPTY和循环体(代码,标签头)
	Private Function GetCloseBlock(ByVal aryTags)
		Dim ary(1)
		If Len(aryTags(0)) > 0 Then
			Dim Match, strSubPattern
			strSubPattern = "\{((?:empty|null|eof|nodata)\:"&aryTags(1)&")\s*?(?:[\s\S.]*?)\/?\}(?:([\s\S.]*?)\{/\1\})"
			Set Match = System.Text.MatchX(aryTags(0), strSubPattern)
			If Match.Count Then'如果有 empty 标签
				ary(0) = Match(0).SubMatches(1) 'empty标签
				ary(1) = System.Text.ReplaceX(aryTags(0), strSubPattern, "") '循环体部分
			Else ary(1) = aryTags(0) End If
			Set Match = Nothing
		End If
		GetCloseBlock = ary
	End Function
	
	'// 获得属性列表,返回名值的字典对象
	Private Function GetBlockAttr(ByVal val)
		Dim i
		Dim Match,Matches,dicAttr
		Set dicAttr = Dicary()'定义字段对象
		'返回一个标签节点的所有属性
		If TypeName(val) = "IXMLDOMElement" Then
			For i = 0 To val.attributes.Length - 1
				dicAttr(val.attributes(i).nodeName) = val.attributes(i).nodeTypedValue
			Next
		Else'存储名值对象
			Set Matches = System.Text.MatchX(val ,"([a-zA-Z0-9]+)\s*=\s*(['|""])([\s\S.]*?)\2")
			For Each Match in Matches'0=属性,2=属性值
				dicAttr(LCase(Trim(Match.SubMatches(0)))) = Match.SubMatches(2)
			Next
			Set Matches = Nothing
		End If
		Set GetBlockAttr = dicAttr
		Set dicAttr = Nothing
	End Function
	
	'// 选择一个带路径的节点,返回解析分解后的路径
	Private Function SelectLabelNode(ByVal strPath)
		Dim Match,Matches : strPath = LCase(Trim(strPath))'标签转换成小写
		Set Matches = System.Text.MatchX( strPath ,"([a-zA-Z0-9\/]+)(\[@?((\w+)=(.*?))?\])?\.?(\w+)?(\:(body|empty|html|null|eof))?")
		'传入参数示例：tag[attr=2].attr2:body
		If Matches.Count Then
			Dim ary(5)
			ary(0) = Matches(0).SubMatches(0) 'tag
			ary(1) = Matches(0).SubMatches(2) 'attr=2
			ary(2) = Matches(0).SubMatches(5) 'attr2
			ary(3) = Matches(0).SubMatches(6) ':body|:empty|:html
			'指定辅助路径
			Dim nodesPath: nodesPath = strRootXMLNode & "/" & ary(0)
			If Len(Matches(0).SubMatches(1)) > 4 Then
				nodesPath = nodesPath & "[@" &ary(1)& "]"
			End If
			
			ary(4) = nodesPath'选择的路径
			SelectLabelNode = ary
		Else SelectLabelNode = Null End If
		Set Matches = Nothing
	End Function
	
	'// 设置节点属性
	Private Function SetLabelAttr(ByVal strPath, ByVal strVal)
		Dim ary,node,i,ii
		strPath = Split(strPath, ",")
		For ii = 0 To Ubound(strPath)
			ary = SelectLabelNode(strPath(ii))'选择标签节点
			If IsArray(ary) = False Then Exit Function
			Select Case LCase(ary(3))
			Case ":body"
				Set node = tplXML.selectNodes(ary(4) & "/body")
					For i = 0 to node.Length - 1
						node(i).childNodes(0).nodeValue = strVal
					Next
			Case ":empty",":null",":eof"
				Set node = tplXML.selectNodes(ary(4) & "/null")
					For i = 0 to node.Length - 1
						node(i).childNodes(0).nodeValue = strVal
					Next
			Case ":html"
				Set node = tplXML.selectNodes(ary(4) & "/html")
					For i = 0 to node.Length - 1
						node(i).childNodes(0).nodeValue = strVal
					Next
			Case Else
				If Len(ary(2)) Then
					Set node = tplXML.selectNodes(ary(4))
					For i = 0 to node.Length - 1
						node(i).setAttribute ary(2),strVal
					Next
				End If
			End Select
		Next
	End Function

	'// 输出结果输出模板的绝对路径
	Private Function AbsPath(ByVal strCode)
		Dim html: html = strCode
		Dim Matches, Match
		Set Matches = System.Text.MatchX(html, "(?:href|src)=(['""|])(?!(\/|\{|\(|\.\/|http:\/\/|https:\/\/|javascript:|#))(.+?)\1")
		For Each Match In Matches
			html = Replace(html, Match.Value, Replace(Match.Value, Match.SubMatches(2), RelPath(Match.SubMatches(2)), 1, -1, 0), 1, -1, 0)
		Next
		Set Matches = Nothing
		AbsPath = html
	End Function
		
	'// 替换相对路径，根据模板路径把../逐层替换到对应的目录
	Private Function RelPath(ByVal strPath)
		strPath = System.Text.ReplaceX(strPath, "(\/|\\)+", "/")
		Dim bRootPath: bRootPath = System.Text.ReplaceX(strRootPath&strTemplatePath, "(\/|\\)+", "/")
		Dim Matches: Set Matches = System.Text.MatchX(strPath, "^(\.\.\/)+")
		
		'// 如果不存在 ../ 父路径
		If Matches.Count = 0 Then
			RelPath = LCase(bRootPath & strPath)
		Else
			Dim src, spath
			'// 模板的全路径
			spath = Split(bRootPath, "/")
			'//看有多少个 ../
			Dim N: N = System.Text.MatchX(Matches(0).Value, "(\.\.\/)").Count
			Dim I: For I=0 To Ubound(spath)-1-N: src = src & spath(I) & "/": Next
			'// 把../替换成正确的目录
			RelPath = LCase(Replace(strPath, Matches(0).Value, src, 1, -1, 0))
		End If
		Set Matches = Nothing
	End Function
End Class
%>