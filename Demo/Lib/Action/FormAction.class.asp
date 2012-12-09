<%
'// 本类由系统自动生成，仅供测试用途
With System.Template
	'// 载入页面
	'.setCache = "boyletpl,4,60"
	'.Root = .Root & "/Index/"
	.File(.Root&"/Index/") = "index.html"

	'// 设置数据库表的前缀
	C("DB.PREFIX") = "BE_"

	'// 对模板中的所有普通标签赋值
	'.d("$") = System.Data.Command("SELECT CATE_NAME AS CATENAME FROM [BE_CATEGORY] WHERE ID=?", Array(Array("id",3,1,4,2)))

	.d("title") = "欢迎使用Boyle.ACL框架 - "

	'// sql="select top 10 id,ci_name as name from be_customer"
	'// 直接在模板中指定SQL语句，将自动执行下面的操作，这样做不安全
	Dim Rs: Set Rs = System.Data.Read(System.Data.ToSQL(Array("BE_CUSTOMER", "ID,CI_NAME,CI_TELEPHONE,CI_ADDRESS", 3), "", ""))
	.d("customer[name=block1]") = Rs
	Set Rs = Nothing

	Dim Customer: Set Customer = M("CUSTOMER")
	Customer.Parameters("") = Array("LIMIT:3", "FIELD:ID,CI_NAME AS NAME,CI_TELEPHONE AS TEL,CI_ADDRESS AS ADDR")
	'Customer.Parameters("FIELD") = Array("ID", "CI_NAME", "CI_TELEPHONE", "CI_ADDRESS")
	Customer.Parameters("WHERE") = Array("id>100 and ci_address<>'-1'", "CI_TELEPHONE='-1'", "_logic:OR")
	Customer.Parameters("ORDER") = Array("ID DESC")
	'Customer.Parameters("") = 10 '// 相当于Customer.Parameters("LIMIT") = 10
	'Customer.Parameters("FIELD") = "ID,CI_NAME AS NAME,CI_TELEPHONE,CI_ADDRESS AS ADDRESS"
	.d("customer[name=block2]") = Customer.Select()
	'.d("sql") = Customer.Parameters("SQL")
	Set Customer = Nothing

	'// ==============================删除添加修改记录示例==============================
	Dim Sale: Set Sale = M("USER")
	'// 多标签批量替换
	'.d("$") = Array(Rs, "demo1,demo2,demo3,title")
	'.d("$") = System.Text.ToHashTable(Array("demo1:数据1", "demo2:数据2", "demo3:数据3", "title:这里是对标题进行更改"))
	.d("$") = Array(Array(Sale.Min("id"), Sale.Sum("id"), Sale.Avg("id"), .GetLabVal("title")&"模板示例"), "demo1,demo2,demo3,title")
	'System.WB Sale.setInc("Cate_Status")
	'System.WB Sale.setDec("Cate_Status:2")
	'Sale.Delete("id=28")
	Sale.Parameters("") = Empty '// 清空所有参数
	'Sale.Add( Array(System.Text.ToHashTable( Array("uName:10000", "uPass:123456", "uDate:"&Now()) )) )
	Dim R1: R1 = Array(Array("uName", "10000"), Array("uPass", "123456"), Array("uDate", Now()))
	Dim R2: R2 = Array(Array("uName", "10001"), Array("uPass", "abcdefg"), Array("uDate", Now()))
	'Sale.Add( Array(R1, R2) )
	'Sale.Add( R1 )
	Sale.Parameters("") = Empty
	Sale.Parameters("WHERE") = "uName='10000'"
	Dim RC: RC = Sale.Count("ID")
	'Sale.Save( R2 )
	.d("sql") = "[" & RC & "]" & Sale.Parameters("SQL")
	Set Sale = Nothing

	'// ==============================分页示例==============================
	Dim blPage: blPage = System.Get("PAGE", 0)
	Dim Parts: Set Parts = M("PARTS")
	Parts.Parameters("") = Array("CURRENTPAGE:"&blPage&"", "PAGESIZE:15", "FIELD:ID,CP_NAME,CP_LOCALITY,CP_CAR")
	Dim PagerResult: PagerResult = Parts.Pager()
	.d("parts") = Array(PagerResult(0), "id,name,locality,car")
	.d("pager") = PagerResult(1)
	'.d("sql") = PagerResult(2)("SQL")
	Set Parts = Nothing

	'// 输出页面
	.Display()

	System.Data.C(Rs)
End With

Call Terminate()

'// 在模板中给块增加dr=function(param)属性来自由设置输出的格式
'// 自定义字段数据输出
Public Function callBack(ByVal blRs)
	'在这个函数里你可以重新定义你的字段数据，就像下面
	blRs("id") = System.Text.AppendZero(blRs("id"), 4)
	'还可以加入你自己定义的名称
	blRs("test") = "这里用做测试"
	'最后用SET方法把数据返回就可以了
	Set callBack = blRs
End Function
%>