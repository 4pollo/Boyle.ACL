<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [系统数据安全操作类]
'// +--------------------------------------------------------------------------
'// | Copyright (c) 2008-2012 By Boyle. All rights reserved.
'// +--------------------------------------------------------------------------
'// | Licensed ( http://www.gnu.org/licenses/gpl.html )
'// +--------------------------------------------------------------------------
'// | Author: Boyle <boyle7[at]qq.com>
'// +--------------------------------------------------------------------------

Class Cls_Security
	
	'// 定义私有命名对象
	Private Private_SHA256, Private_MD5, Private_AES
	
	'// 初始化类
	Private Sub Class_Initialize()
	End Sub	
	
	'// 释放类
	Private Sub Class_Terminate()
		If IsObject(Private_AES) Then Set Private_AES = Nothing
		If IsObject(Private_MD5) Then Set Private_MD5 = Nothing
		If IsObject(Private_SHA256) Then Set Private_SHA256 = Nothing
	End Sub
	
	'// 声明对象模块单元
	Public Property Get AES()
		If Not IsObject(Private_AES) Then Set Private_AES = New Cls_Security_AES End If
		Set AES = Private_AES
	End Property
	Public Property Get MD5(ByVal strVal, ByVal numVal)
		If Not IsOBject(Private_MD5) Then Set Private_MD5 = New Cls_Security_MD5 End If
		MD5 = Private_MD5.Encrypt(strVal, numVal)
	End Property
	Public Property Get SHA256(ByVal strVal)
		If Not IsObject(Private_SHA256) Then Set Private_SHA256 = New Cls_Security_SHA256 End If
		SHA256 = Private_SHA256.Encrypt(strVal)
	End Property
End Class
%>

<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [MD5加密类]
'// +--------------------------------------------------------------------------
Class Cls_Security_MD5
	Private BITS_TO_A_BYTE
	Private BYTES_TO_A_WORD
	Private BITS_TO_A_WORD

	Private m_lOnBits(30)
	Private m_l2Power(30)

	Dim Md5OLD

	Private Sub Class_Initialize()
		BITS_TO_A_BYTE = 8
		BYTES_TO_A_WORD = 4
		BITS_TO_A_WORD = 32
	End Sub
	Private Sub Class_Terminate()
	End Sub

	Private Function LShift(lValue, iShiftBits)
		If iShiftBits = 0 Then
			LShift = lValue
			Exit Function
		ElseIf iShiftBits = 31 Then
			If lValue And 1 Then
				LShift = &H80000000
			Else
				LShift = 0
			End If
			Exit Function
		ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
			Err.Raise 6
		End If

		If (lValue And m_l2Power(31 - iShiftBits)) Then
			LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
		Else
			LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
		End If
	End Function

	Private Function str2bin(varstr)
		Dim varasc
		Dim i
		Dim varchar
		Dim varlow
		Dim varhigh
		
		str2bin = ""
		For i = 1 To Len(varstr)
			varchar = Mid(varstr, i, 1)
			varasc = Asc(varchar)
			
			If varasc < 0 Then
			varasc = varasc + 65535
			End If
			
			If varasc > 255 Then
			varlow = Left(Hex(Asc(varchar)), 2)
			varhigh = Right(Hex(Asc(varchar)), 2)
			str2bin = str2bin & ChrB("&H" & varlow) & ChrB("&H" & varhigh)
			Else
			str2bin = str2bin & ChrB(AscB(varchar))
			End If
		Next
	End Function

	Private Function RShift(lValue, iShiftBits)
		If iShiftBits = 0 Then
			RShift = lValue
			Exit Function
		ElseIf iShiftBits = 31 Then
			If lValue And &H80000000 Then
				RShift = 1
			Else
				RShift = 0
			End If
			Exit Function
		ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
			Err.Raise 6
		End If

		RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

		If (lValue And &H80000000) Then
			RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
		End If
	End Function

	Private Function RotateLeft(lValue, iShiftBits)
		RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
	End Function

	Private Function AddUnsigned(lX, lY)
		Dim lX4
		Dim lY4
		Dim lX8
		Dim lY8
		Dim lResult

		lX8 = lX And &H80000000
		lY8 = lY And &H80000000
		lX4 = lX And &H40000000
		lY4 = lY And &H40000000
		
		lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

		If lX4 And lY4 Then
			lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
		ElseIf lX4 Or lY4 Then
			If lResult And &H40000000 Then
				lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
			Else
				lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
			End If
		Else
			lResult = lResult Xor lX8 Xor lY8
		End If

		AddUnsigned = lResult
	End Function

	Private Function md5_F(x, y, z)
		md5_F = (x And y) Or ((Not x) And z)
	End Function

	Private Function md5_G(x, y, z)
		md5_G = (x And z) Or (y And (Not z))
	End Function

	Private Function md5_H(x, y, z)
		md5_H = (x Xor y Xor z)
	End Function

	Private Function md5_I(x, y, z)
		md5_I = (y Xor (x Or (Not z)))
	End Function

	Private Sub md5_FF(a, b, c, d, x, s, ac)
		a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_F(b, c, d), x), ac))
		a = RotateLeft(a, s)
		a = AddUnsigned(a, b)
	End Sub

	Private Sub md5_GG(a, b, c, d, x, s, ac)
		a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_G(b, c, d), x), ac))
		a = RotateLeft(a, s)
		a = AddUnsigned(a, b)
	End Sub

	Private Sub md5_HH(a, b, c, d, x, s, ac)
		a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_H(b, c, d), x), ac))
		a = RotateLeft(a, s)
		a = AddUnsigned(a, b)
	End Sub

	Private Sub md5_II(a, b, c, d, x, s, ac)
		a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_I(b, c, d), x), ac))
		a = RotateLeft(a, s)
		a = AddUnsigned(a, b)
	End Sub

	Private Function ConvertToWordArray(sMessage)
		Dim lMessageLength
		Dim lNumberOfWords
		Dim lWordArray()
		Dim lBytePosition
		Dim lByteCount
		Dim lWordCount
		
		Const MODULUS_BITS = 512
		Const CONGRUENT_BITS = 448
		If Md5OLD = 1 Then
			lMessageLength = Len(sMessage)
		Else
			lMessageLength = LenB(sMessage)
		End If
		lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
		ReDim lWordArray(lNumberOfWords - 1)
		
		lBytePosition = 0
		lByteCount = 0
		Do Until lByteCount >= lMessageLength
			lWordCount = lByteCount \ BYTES_TO_A_WORD
			lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
			lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(AscB(MidB(sMessage, lByteCount + 1, 1)), lBytePosition)
			lByteCount = lByteCount + 1
		Loop

		lWordCount = lByteCount \ BYTES_TO_A_WORD
		lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
		
		lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
		
		lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
		lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)

		ConvertToWordArray = lWordArray
	End Function

	Private Function WordToHex(lValue)
		Dim lByte
		Dim lCount
		
		For lCount = 0 To 3
			lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
			WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
		Next
	End Function

	Public Function Encrypt(sMessage, sType)
		m_lOnBits(0) = CLng(1)
		m_lOnBits(1) = CLng(3)
		m_lOnBits(2) = CLng(7)
		m_lOnBits(3) = CLng(15)
		m_lOnBits(4) = CLng(31)
		m_lOnBits(5) = CLng(63)
		m_lOnBits(6) = CLng(127)
		m_lOnBits(7) = CLng(255)
		m_lOnBits(8) = CLng(511)
		m_lOnBits(9) = CLng(1023)
		m_lOnBits(10) = CLng(2047)
		m_lOnBits(11) = CLng(4095)
		m_lOnBits(12) = CLng(8191)
		m_lOnBits(13) = CLng(16383)
		m_lOnBits(14) = CLng(32767)
		m_lOnBits(15) = CLng(65535)
		m_lOnBits(16) = CLng(131071)
		m_lOnBits(17) = CLng(262143)
		m_lOnBits(18) = CLng(524287)
		m_lOnBits(19) = CLng(1048575)
		m_lOnBits(20) = CLng(2097151)
		m_lOnBits(21) = CLng(4194303)
		m_lOnBits(22) = CLng(8388607)
		m_lOnBits(23) = CLng(16777215)
		m_lOnBits(24) = CLng(33554431)
		m_lOnBits(25) = CLng(67108863)
		m_lOnBits(26) = CLng(134217727)
		m_lOnBits(27) = CLng(268435455)
		m_lOnBits(28) = CLng(536870911)
		m_lOnBits(29) = CLng(1073741823)
		m_lOnBits(30) = CLng(2147483647)
		
		m_l2Power(0) = CLng(1)
		m_l2Power(1) = CLng(2)
		m_l2Power(2) = CLng(4)
		m_l2Power(3) = CLng(8)
		m_l2Power(4) = CLng(16)
		m_l2Power(5) = CLng(32)
		m_l2Power(6) = CLng(64)
		m_l2Power(7) = CLng(128)
		m_l2Power(8) = CLng(256)
		m_l2Power(9) = CLng(512)
		m_l2Power(10) = CLng(1024)
		m_l2Power(11) = CLng(2048)
		m_l2Power(12) = CLng(4096)
		m_l2Power(13) = CLng(8192)
		m_l2Power(14) = CLng(16384)
		m_l2Power(15) = CLng(32768)
		m_l2Power(16) = CLng(65536)
		m_l2Power(17) = CLng(131072)
		m_l2Power(18) = CLng(262144)
		m_l2Power(19) = CLng(524288)
		m_l2Power(20) = CLng(1048576)
		m_l2Power(21) = CLng(2097152)
		m_l2Power(22) = CLng(4194304)
		m_l2Power(23) = CLng(8388608)
		m_l2Power(24) = CLng(16777216)
		m_l2Power(25) = CLng(33554432)
		m_l2Power(26) = CLng(67108864)
		m_l2Power(27) = CLng(134217728)
		m_l2Power(28) = CLng(268435456)
		m_l2Power(29) = CLng(536870912)
		m_l2Power(30) = CLng(1073741824)
		
		
		Dim x
		Dim k
		Dim AA
		Dim BB
		Dim CC
		Dim DD
		Dim a
		Dim b
		Dim c
		Dim d
		
		Const S11 = 7
		Const S12 = 12
		Const S13 = 17
		Const S14 = 22
		Const S21 = 5
		Const S22 = 9
		Const S23 = 14
		Const S24 = 20
		Const S31 = 4
		Const S32 = 11
		Const S33 = 16
		Const S34 = 23
		Const S41 = 6
		Const S42 = 10
		Const S43 = 15
		Const S44 = 21
		If Md5OLD = 1 Then
			x = ConvertToWordArray(sMessage)
		Else
			x = ConvertToWordArray(str2bin(sMessage))
		End If
		a = &H67452301
		b = &HEFCDAB89
		c = &H98BADCFE
		d = &H10325476
		
		For k = 0 To UBound(x) Step 16
			AA = a
			BB = b
			CC = c
			DD = d
			
			md5_FF a, b, c, d, x(k + 0), S11, &HD76AA478
			md5_FF d, a, b, c, x(k + 1), S12, &HE8C7B756
			md5_FF c, d, a, b, x(k + 2), S13, &H242070DB
			md5_FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
			md5_FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
			md5_FF d, a, b, c, x(k + 5), S12, &H4787C62A
			md5_FF c, d, a, b, x(k + 6), S13, &HA8304613
			md5_FF b, c, d, a, x(k + 7), S14, &HFD469501
			md5_FF a, b, c, d, x(k + 8), S11, &H698098D8
			md5_FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
			md5_FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
			md5_FF b, c, d, a, x(k + 11), S14, &H895CD7BE
			md5_FF a, b, c, d, x(k + 12), S11, &H6B901122
			md5_FF d, a, b, c, x(k + 13), S12, &HFD987193
			md5_FF c, d, a, b, x(k + 14), S13, &HA679438E
			md5_FF b, c, d, a, x(k + 15), S14, &H49B40821
			
			md5_GG a, b, c, d, x(k + 1), S21, &HF61E2562
			md5_GG d, a, b, c, x(k + 6), S22, &HC040B340
			md5_GG c, d, a, b, x(k + 11), S23, &H265E5A51
			md5_GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
			md5_GG a, b, c, d, x(k + 5), S21, &HD62F105D
			md5_GG d, a, b, c, x(k + 10), S22, &H2441453
			md5_GG c, d, a, b, x(k + 15), S23, &HD8A1E681
			md5_GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
			md5_GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
			md5_GG d, a, b, c, x(k + 14), S22, &HC33707D6
			md5_GG c, d, a, b, x(k + 3), S23, &HF4D50D87
			md5_GG b, c, d, a, x(k + 8), S24, &H455A14ED
			md5_GG a, b, c, d, x(k + 13), S21, &HA9E3E905
			md5_GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
			md5_GG c, d, a, b, x(k + 7), S23, &H676F02D9
			md5_GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
			
			md5_HH a, b, c, d, x(k + 5), S31, &HFFFA3942
			md5_HH d, a, b, c, x(k + 8), S32, &H8771F681
			md5_HH c, d, a, b, x(k + 11), S33, &H6D9D6122
			md5_HH b, c, d, a, x(k + 14), S34, &HFDE5380C
			md5_HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
			md5_HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
			md5_HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
			md5_HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
			md5_HH a, b, c, d, x(k + 13), S31, &H289B7EC6
			md5_HH d, a, b, c, x(k + 0), S32, &HEAA127FA
			md5_HH c, d, a, b, x(k + 3), S33, &HD4EF3085
			md5_HH b, c, d, a, x(k + 6), S34, &H4881D05
			md5_HH a, b, c, d, x(k + 9), S31, &HD9D4D039
			md5_HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
			md5_HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
			md5_HH b, c, d, a, x(k + 2), S34, &HC4AC5665
			
			md5_II a, b, c, d, x(k + 0), S41, &HF4292244
			md5_II d, a, b, c, x(k + 7), S42, &H432AFF97
			md5_II c, d, a, b, x(k + 14), S43, &HAB9423A7
			md5_II b, c, d, a, x(k + 5), S44, &HFC93A039
			md5_II a, b, c, d, x(k + 12), S41, &H655B59C3
			md5_II d, a, b, c, x(k + 3), S42, &H8F0CCC92
			md5_II c, d, a, b, x(k + 10), S43, &HFFEFF47D
			md5_II b, c, d, a, x(k + 1), S44, &H85845DD1
			md5_II a, b, c, d, x(k + 8), S41, &H6FA87E4F
			md5_II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
			md5_II c, d, a, b, x(k + 6), S43, &HA3014314
			md5_II b, c, d, a, x(k + 13), S44, &H4E0811A1
			md5_II a, b, c, d, x(k + 4), S41, &HF7537E82
			md5_II d, a, b, c, x(k + 11), S42, &HBD3AF235
			md5_II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
			md5_II b, c, d, a, x(k + 9), S44, &HEB86D391
			
			a = AddUnsigned(a, AA)
			b = AddUnsigned(b, BB)
			c = AddUnsigned(c, CC)
			d = AddUnsigned(d, DD)
		Next
		
		If sType = 32 Then
		    Encrypt = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
		Else
		    Encrypt = LCase(WordToHex(b) & WordToHex(c)) 'I crop this to fit 16byte database password :D
		End If
	End Function
End Class
%>

<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [SHA256加密类]
'// +--------------------------------------------------------------------------
Class Cls_Security_SHA256
	Private m_lOnBits(30)
	Private m_l2Power(30)
	Private K(63)

	Private BITS_TO_A_BYTE
	Private BYTES_TO_A_WORD
	Private BITS_TO_A_WORD

	Private Sub Class_Initialize()
		BITS_TO_A_BYTE  = 8
		BYTES_TO_A_WORD = 4
		BITS_TO_A_WORD  = 32
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Private Function LShift(lValue, iShiftBits)
		If iShiftBits = 0 Then
			LShift = lValue
			Exit Function
		ElseIf iShiftBits = 31 Then
			If lValue And 1 Then
				LShift = &H80000000
			Else
				LShift = 0
			End If
			Exit Function
		ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
			Err.Raise 6
		End If
		
		If (lValue And m_l2Power(31 - iShiftBits)) Then
			LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
		Else
			LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
		End If
	End Function

	Private Function RShift(lValue, iShiftBits)
		If iShiftBits = 0 Then
			RShift = lValue
			Exit Function
		ElseIf iShiftBits = 31 Then
			If lValue And &H80000000 Then
				RShift = 1
			Else
				RShift = 0
			End If
			Exit Function
		ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
			Err.Raise 6
		End If
		
		RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
		
		If (lValue And &H80000000) Then
			RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
		End If
	End Function

	Private Function AddUnsigned(lX, lY)
		Dim lX4
		Dim lY4
		Dim lX8
		Dim lY8
		Dim lResult
	 
		lX8 = lX And &H80000000
		lY8 = lY And &H80000000
		lX4 = lX And &H40000000
		lY4 = lY And &H40000000
	 
		lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
	 
		If lX4 And lY4 Then
			lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
		ElseIf lX4 Or lY4 Then
			If lResult And &H40000000 Then
				lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
			Else
				lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
			End If
		Else
			lResult = lResult Xor lX8 Xor lY8
		End If
	 
		AddUnsigned = lResult
	End Function

	Private Function Ch(x, y, z)
		Ch = ((x And y) Xor ((Not x) And z))
	End Function

	Private Function Maj(x, y, z)
		Maj = ((x And y) Xor (x And z) Xor (y And z))
	End Function

	Private Function S(x, n)
		S = (RShift(x, (n And m_lOnBits(4))) Or LShift(x, (32 - (n And m_lOnBits(4)))))
	End Function

	Private Function R(x, n)
		R = RShift(x, CInt(n And m_lOnBits(4)))
	End Function

	Private Function Sigma0(x)
		Sigma0 = (S(x, 2) Xor S(x, 13) Xor S(x, 22))
	End Function

	Private Function Sigma1(x)
		Sigma1 = (S(x, 6) Xor S(x, 11) Xor S(x, 25))
	End Function

	Private Function Gamma0(x)
		Gamma0 = (S(x, 7) Xor S(x, 18) Xor R(x, 3))
	End Function

	Private Function Gamma1(x)
		Gamma1 = (S(x, 17) Xor S(x, 19) Xor R(x, 10))
	End Function

	Private Function ConvertToWordArray(sMessage)
		Dim lMessageLength
		Dim lNumberOfWords
		Dim lWordArray()
		Dim lBytePosition
		Dim lByteCount
		Dim lWordCount
		Dim lByte
		
		Const MODULUS_BITS = 512
		Const CONGRUENT_BITS = 448
		
		lMessageLength = Len(sMessage)
		
		lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
		ReDim lWordArray(lNumberOfWords - 1)
		
		lBytePosition = 0
		lByteCount = 0
		Do Until lByteCount >= lMessageLength
			lWordCount = lByteCount \ BYTES_TO_A_WORD
			
			lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
			
			lByte = AscB(Mid(sMessage, lByteCount + 1, 1))
			
			lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
			lByteCount = lByteCount + 1
		Loop

		lWordCount = lByteCount \ BYTES_TO_A_WORD
		lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE

		lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

		lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
		lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)
		
		ConvertToWordArray = lWordArray
	End Function

	Public Function Encrypt(sMessage)
		m_lOnBits(0) = CLng(1)
		m_lOnBits(1) = CLng(3)
		m_lOnBits(2) = CLng(7)
		m_lOnBits(3) = CLng(15)
		m_lOnBits(4) = CLng(31)
		m_lOnBits(5) = CLng(63)
		m_lOnBits(6) = CLng(127)
		m_lOnBits(7) = CLng(255)
		m_lOnBits(8) = CLng(511)
		m_lOnBits(9) = CLng(1023)
		m_lOnBits(10) = CLng(2047)
		m_lOnBits(11) = CLng(4095)
		m_lOnBits(12) = CLng(8191)
		m_lOnBits(13) = CLng(16383)
		m_lOnBits(14) = CLng(32767)
		m_lOnBits(15) = CLng(65535)
		m_lOnBits(16) = CLng(131071)
		m_lOnBits(17) = CLng(262143)
		m_lOnBits(18) = CLng(524287)
		m_lOnBits(19) = CLng(1048575)
		m_lOnBits(20) = CLng(2097151)
		m_lOnBits(21) = CLng(4194303)
		m_lOnBits(22) = CLng(8388607)
		m_lOnBits(23) = CLng(16777215)
		m_lOnBits(24) = CLng(33554431)
		m_lOnBits(25) = CLng(67108863)
		m_lOnBits(26) = CLng(134217727)
		m_lOnBits(27) = CLng(268435455)
		m_lOnBits(28) = CLng(536870911)
		m_lOnBits(29) = CLng(1073741823)
		m_lOnBits(30) = CLng(2147483647)

		m_l2Power(0) = CLng(1)
		m_l2Power(1) = CLng(2)
		m_l2Power(2) = CLng(4)
		m_l2Power(3) = CLng(8)
		m_l2Power(4) = CLng(16)
		m_l2Power(5) = CLng(32)
		m_l2Power(6) = CLng(64)
		m_l2Power(7) = CLng(128)
		m_l2Power(8) = CLng(256)
		m_l2Power(9) = CLng(512)
		m_l2Power(10) = CLng(1024)
		m_l2Power(11) = CLng(2048)
		m_l2Power(12) = CLng(4096)
		m_l2Power(13) = CLng(8192)
		m_l2Power(14) = CLng(16384)
		m_l2Power(15) = CLng(32768)
		m_l2Power(16) = CLng(65536)
		m_l2Power(17) = CLng(131072)
		m_l2Power(18) = CLng(262144)
		m_l2Power(19) = CLng(524288)
		m_l2Power(20) = CLng(1048576)
		m_l2Power(21) = CLng(2097152)
		m_l2Power(22) = CLng(4194304)
		m_l2Power(23) = CLng(8388608)
		m_l2Power(24) = CLng(16777216)
		m_l2Power(25) = CLng(33554432)
		m_l2Power(26) = CLng(67108864)
		m_l2Power(27) = CLng(134217728)
		m_l2Power(28) = CLng(268435456)
		m_l2Power(29) = CLng(536870912)
		m_l2Power(30) = CLng(1073741824)
			
		K(0) = &H428A2F98
		K(1) = &H71374491
		K(2) = &HB5C0FBCF
		K(3) = &HE9B5DBA5
		K(4) = &H3956C25B
		K(5) = &H59F111F1
		K(6) = &H923F82A4
		K(7) = &HAB1C5ED5
		K(8) = &HD807AA98
		K(9) = &H12835B01
		K(10) = &H243185BE
		K(11) = &H550C7DC3
		K(12) = &H72BE5D74
		K(13) = &H80DEB1FE
		K(14) = &H9BDC06A7
		K(15) = &HC19BF174
		K(16) = &HE49B69C1
		K(17) = &HEFBE4786
		K(18) = &HFC19DC6
		K(19) = &H240CA1CC
		K(20) = &H2DE92C6F
		K(21) = &H4A7484AA
		K(22) = &H5CB0A9DC
		K(23) = &H76F988DA
		K(24) = &H983E5152
		K(25) = &HA831C66D
		K(26) = &HB00327C8
		K(27) = &HBF597FC7
		K(28) = &HC6E00BF3
		K(29) = &HD5A79147
		K(30) = &H6CA6351
		K(31) = &H14292967
		K(32) = &H27B70A85
		K(33) = &H2E1B2138
		K(34) = &H4D2C6DFC
		K(35) = &H53380D13
		K(36) = &H650A7354
		K(37) = &H766A0ABB
		K(38) = &H81C2C92E
		K(39) = &H92722C85
		K(40) = &HA2BFE8A1
		K(41) = &HA81A664B
		K(42) = &HC24B8B70
		K(43) = &HC76C51A3
		K(44) = &HD192E819
		K(45) = &HD6990624
		K(46) = &HF40E3585
		K(47) = &H106AA070
		K(48) = &H19A4C116
		K(49) = &H1E376C08
		K(50) = &H2748774C
		K(51) = &H34B0BCB5
		K(52) = &H391C0CB3
		K(53) = &H4ED8AA4A
		K(54) = &H5B9CCA4F
		K(55) = &H682E6FF3
		K(56) = &H748F82EE
		K(57) = &H78A5636F
		K(58) = &H84C87814
		K(59) = &H8CC70208
		K(60) = &H90BEFFFA
		K(61) = &HA4506CEB
		K(62) = &HBEF9A3F7
		K(63) = &HC67178F2

		Dim HASH(7)
		Dim M
		Dim W(63)
		Dim a
		Dim b
		Dim c
		Dim d
		Dim e
		Dim f
		Dim g
		Dim h
		Dim i
		Dim j
		Dim T1
		Dim T2
		
		HASH(0) = &H6A09E667
		HASH(1) = &HBB67AE85
		HASH(2) = &H3C6EF372
		HASH(3) = &HA54FF53A
		HASH(4) = &H510E527F
		HASH(5) = &H9B05688C
		HASH(6) = &H1F83D9AB
		HASH(7) = &H5BE0CD19
		
		M = ConvertToWordArray(sMessage)
		
		For i = 0 To UBound(M) Step 16
			a = HASH(0)
			b = HASH(1)
			c = HASH(2)
			d = HASH(3)
			e = HASH(4)
			f = HASH(5)
			g = HASH(6)
			h = HASH(7)
			
			For j = 0 To 63
				If j < 16 Then
					W(j) = M(j + i)
				Else
					W(j) = AddUnsigned(AddUnsigned(AddUnsigned(Gamma1(W(j - 2)), W(j - 7)), Gamma0(W(j - 15))), W(j - 16))
				End If
					
				T1 = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(h, Sigma1(e)), Ch(e, f, g)), K(j)), W(j))
				T2 = AddUnsigned(Sigma0(a), Maj(a, b, c))
				
				h = g
				g = f
				f = e
				e = AddUnsigned(d, T1)
				d = c
				c = b
				b = a
				a = AddUnsigned(T1, T2)
			Next
			
			HASH(0) = AddUnsigned(a, HASH(0))
			HASH(1) = AddUnsigned(b, HASH(1))
			HASH(2) = AddUnsigned(c, HASH(2))
			HASH(3) = AddUnsigned(d, HASH(3))
			HASH(4) = AddUnsigned(e, HASH(4))
			HASH(5) = AddUnsigned(f, HASH(5))
			HASH(6) = AddUnsigned(g, HASH(6))
			HASH(7) = AddUnsigned(h, HASH(7))
		Next
		
		Encrypt = LCase(Right("00000000" & Hex(HASH(0)), 8) & Right("00000000" & Hex(HASH(1)), 8) & Right("00000000" & Hex(HASH(2)), 8) & Right("00000000" & Hex(HASH(3)), 8) & Right("00000000" & Hex(HASH(4)), 8) & Right("00000000" & Hex(HASH(5)), 8) & Right("00000000" & Hex(HASH(6)), 8) & Right("00000000" & Hex(HASH(7)), 8))
	End Function
End Class
%>

<%
'// +--------------------------------------------------------------------------
'// | Boyle.ACL [AES加密解密类]
'// +--------------------------------------------------------------------------
Class Cls_Security_AES
	Private m_KeySize, m_Key

	Private Sub Class_Initialize()
		'// 初始化必要参数
		m_Key = "BOYLE.ACL": m_KeySize = 128		
		BuildSBox(): BuildIsBox(): BuildRcon()
	End Sub
	
	Private Sub Class_Terminate()
	End Sub
	
	Public Property Get KeySize()
		KeySize = m_KeySize
	End Property
	Public Property Let KeySize(ByVal blParam)
		m_KeySize = blParam
	End Property
	Public Property Get Key()
		Key = m_Key
	End Property
	Public Property Let Key(ByVal blParam)
		m_Key = blParam
	End Property
	
	Public Function Encrypt(ByVal blParam)
		Encrypt = CipherStrToHexStr(blParam)
	End Function	
	Public Function Decrypt(ByVal blParam)
		Decrypt = InvCipherHexStrToStr(blParam)
	End Function
	
	'**********************************************
	' 输入：keySize密钥长度(128、192、256),Src明文字符串，Key密钥字符串
	'       明文字符串不能超过 &HFFFF长度
	' 输出：密文十六进制字符串
	'**********************************************
	Public Function CipherStrToHexStr(ByVal Src)
		SetNbNkNr()
		InitKey()
		Dim I, Result, Str32, Input(15), Output(15)
		Dim HexString, iLen, sLen
		iLen = Len(Src)
		sLen = CStr(Hex(iLen))
		sLen = String(4-Len(sLen), "0") & sLen
		HexString = sLen & HexStr(Src)
		Result = ""
		I = 0
		Str32 = Mid(HexString, 1, 32)
		Do While Len(Str32) > 0
			HexStrToArray Str32, Input
			FCipher Input, Output
			Result = Result + ArrayToHexStr(Output)
			I = I + Len(Str32)
			Str32 = Mid(HexString, I + 1, 32)
		Loop
		CipherStrToHexStr = Result
	End Function
	
	'**********************************************
	' 输入：keySize密钥长度(128、192、256),Src明文十六进制符串，Key密钥字符串
	'       明文十六进制字符串不能超过 2 * &HFFFF长度
	' 输出：密文十六进制字符串
	'**********************************************	
	Public Function CipherHexStrToHexStr(ByVal Src)
		SetNbNkNr()
		InitKey()
		Dim I, Result, Str32, Input(15), Output(15)
		Dim HexString, iLen, sLen
		iLen = Len(Src) \ 2
		if iLen > 2 * &HFFFF then Src = Left(Src, 2 * &HFFFF)
		sLen = CStr(Hex(iLen))
		sLen = String(4-Len(sLen), "0")&sLen
		HexString = sLen & Src
		Result = ""
		I = 0
		Str32 = Mid(HexString, 1, 32)
		Do While Len(Str32) > 0
			HexStrToArray Str32, Input
			FCipher Input, Output
			Result = Result + ArrayToHexStr(Output)
			I = I + Len(Str32)
			Str32 = Mid(HexString, I + 1, 32)
		Loop
		CipherHexStrToHexStr = Result
	End Function
	
	'**********************************************
	' 输入：keySize密钥长度(128、192、256),Src密文十六进制符串，Key密钥字符串
	' 输出：解密后的字符串
	'**********************************************	
	Public Function InvCipherHexStrToStr(ByVal Src)
		SetNbNkNr()
		InitKey()
		Dim I, Result, Str32, Input(15), Output(15)
		Dim HexString, iLen, sLen, Str
		HexString = Src
		Result = ""
		I = 0
		Str32 = Mid(HexString, 1, 32)
		I = I + Len(Str32)
		HexStrToArray Str32, Input
		FInvCipher Input, Output
		Str = ArrayToHexStr(Output)
		sLen = Left(Str, 4)
		iLen = HexToLng(sLen)
		Str = ArrayToStr(Output)
		Result = Right(Str, 7)
		Str32 = Mid(HexString, I + 1, 32)	
		Do While Len(Str32) > 0
			HexStrToArray Str32, Input
			FInvCipher Input, Output
			Result = Result + ArrayToStr(Output)
			I = I + Len(Str32)
			Str32 = Mid(HexString, I + 1, 32)
		Loop
		InvCipherHexStrToStr = Left(Result, iLen)
	End Function
	
	'**********************************************
	' 输入：keySize密钥长度(128、192、256),Src密文十六进制符串，Key密钥字符串
	' 输出：解密后的十六进制字符串
	'**********************************************
	Public Function InvCipherHexStrToHexStr(ByVal Src)
		SetNbNkNr()
		InitKey()
		Dim I, Result, Str32, Input(15), Output(15)
		Dim HexString, iLen, sLen, Str
		HexString = Src
		Result = ""
		I = 0
		Str32 = Mid(HexString, 1, 32)
		I = I + Len(Str32)
		HexStrToArray Str32, Input
		FInvCipher Input, Output
		Str = ArrayToHexStr(Output)
		sLen = Left(Str, 4)
		iLen = HexToLng(sLen)
		Result = Right(Str, 28)
		Str32 = Mid(HexString, I + 1, 32)	
		Do While Len(Str32) > 0
			HexStrToArray Str32, Input
			FInvCipher Input, Output
			Result = Result + ArrayToHexStr(Output)
			I = I + Len(Str32)
			Str32 = Mid(HexString, I + 1, 32)
		Loop
		InvCipherHexStrToHexStr = Left(Result, iLen * 4)
	End Function
	
	'**********************************************
	' 类的实现
	'**********************************************
	Private FSBox(15, 15)
	Private FIsBox(15, 15)
	Private FRcon(10, 3)
	Private FNb, FNk, FNr
	Private FKey(31)
	Private FW(59, 3)
	Private FState(3, 3)
	
	Private Function ArrayToHexStr(Src)
		Dim I, Result: Result = ""
		For I = LBound(Src) To UBound(Src)
			Result = Result&CStr(MyHex(Src(I)))
		Next
		ArrayToHexStr = Result
	End Function
	
	Private Function ArrayToStr(Src)
		Dim I, Result: Result = ""
		For I = LBound(Src) To UBound(Src) \ 2
			Result = Result&ChrW(Src(2 * I) + Src(2 * I + 1) * &H100)
		Next
		ArrayToStr = Result
	End Function
	
	Private Function HexStr(Src)
		Dim I, HexString
		For I = 0 To LenB(Src) - 1
			HexString = HexString&CStr(MyHex(AscB(MidB(Src, I + 1, 1))))
		Next
		HexStr = HexString
	End Function
	
	Private Function HexToLng(H)
		HexToLng = CLng(Cstr("&H" & H))
	End Function
	
	Private Sub HexStrToArray(Src, Out)
		If IsNull(Src) then Src = ""
		Dim W, I, J: I = 0: J = 0
		For I = 0 To Len(Src) \ 2 - 1
			Out(I) = HexToLng(Mid(Src, 2*I + 1, 2))
		Next
		For I = Len(Src) \ 2 To 15
			Out(I) = 0
		Next
	End Sub
	
	Private Function CByte(B)
		CByte = B And &H00FF
	End Function
	
	Private Function MyHex(B)
		If B < &H10 then MyHex = "0"&CStr(Hex(B)) Else MyHex = CStr(Hex(B)) End If
	End Function
	
	'**********************************************
	' 初始化工作Key，如果Key中包含Unicode 字符，则仅取Unicode字符的低字节
	'**********************************************
	Private Sub InitKey()
		Dim I, J, K
		For I = 0 To 31
			FKey(I) = 0
		Next
		If Len(m_Key) > FNk * 4 then
			For I = 0 To FNk * 4 - 1
				K = AscW(Mid(m_Key, I + 1, 1))
				If K > &HFF then K = CByte(K)
				FKey(I) = K
			Next
		Else
			For I = 0 To len(m_Key) - 1
				K = AscW(Mid(m_Key, I + 1, 1))
				If K > &HFF then K = CByte(K)
				FKey(I) = K
			Next		
		End If
		KeyExpansion
	End Sub	
	
	Private Sub SetNbNkNr()
		FNb = 4
		Select Case m_KeySize
			Case 192: FNk = 6: FNr = 12
			Case 256: FNk = 8: FNr = 14
			'// 否则的都按128 处理
			Case Else FNk = 4: FNr = 10
		End Select
	End Sub
	
	Private Sub AddRoundKey(around)
		Dim R, C: For R = 0 To 3
			For C = 0 To 3
				FState(R, C) = CByte((CLng(FState(R, C)) Xor (Fw((around * 4) + C, R))))
			Next
		Next
	End Sub
	
	Private Sub KeyExpansion()
		Dim Row
		Dim Temp(3)
		Dim I
		For Row = 0 To FNk - 1 
			FW(Row, 0) = FKey(4 * Row)
			FW(Row, 1) = FKey(4 * Row + 1)
			FW(Row, 2) = FKey(4 * Row + 2)
			FW(Row, 3) = FKey(4 * Row + 3)
		Next
		For Row = FNk To FNb * (FNr + 1) - 1
			Temp(0) = FW(Row - 1, 0)
			Temp(1) = FW(Row - 1, 1)
			Temp(2) = FW(Row - 1, 2)
			Temp(3) = FW(Row - 1, 3)
			If Row Mod FNk = 0 then
				RotWord Temp(0), Temp(1), Temp(2), Temp(3)
				SubWord Temp(0), Temp(1), Temp(2), Temp(3)
				Temp(0) = CByte((CLng(Temp(0))) Xor (CLng(FRcon(Row \ FNk, 0))))
				Temp(1) = CByte((CLng(Temp(1))) Xor (CLng(FRcon(Row \ FNk, 1))))
				Temp(2) = CByte((CLng(Temp(2))) Xor (CLng(FRcon(Row \ FNk, 2))))
				Temp(3) = CByte((CLng(Temp(3))) Xor (CLng(FRcon(Row \ FNk, 3))))
			Else 
				If (FNK > 6) And ((Row Mod FNk) = 4) then SubWord Temp(0), Temp(1), Temp(2), Temp(3)
			End If
			FW(Row, 0) = CByte((CLng(FW(Row-FNk, 0))) Xor (CLng(Temp(0))))
			FW(Row, 1) = CByte((CLng(FW(Row-FNk, 1))) Xor (CLng(Temp(1))))
			FW(Row, 2) = CByte((CLng(FW(Row-FNk, 2))) Xor (CLng(Temp(2))))
			FW(Row, 3) = CByte((CLng(FW(Row-FNk, 3))) Xor (CLng(Temp(3))))
		Next
	End Sub
	
	Private Sub SubBytes()
		Dim R, C: For R = 0 To 3
			For C = 0 To 3
				FState(R, C) = FSBox(FState(R, C) \ 16, FState(R, C) And &H0F)
			Next
		Next
	End Sub
	
	Private Sub InvSubBytes()
		Dim R, C: For R = 0 To 3
			For C = 0 To 3
				FState(R, C) = FIsBox(FState(R, C) \ 16, FState(R, C) And &H0F)
			Next
		Next
	End Sub
	
	Private Sub ShIftRows()
		Dim Temp(3, 3)
		Dim R, C
		For R = 0 To 3
			For C = 0 To 3
				Temp(R, C) = FState(R, C)
			Next
		Next
		For R = 1 To 3
			For C = 0 To 3
				FState(R, C) = Temp(R, (C + R) Mod FNb)
			Next
		Next
	End Sub
	
	Private Sub InvShIftRows()
		Dim Temp(3, 3)
		Dim R, C
		For R = 0 To 3
			For C = 0 To 3
				Temp(R, C) = FState(R, C)
			Next
		Next
		For R = 1 To 3
			For C = 0 To 3
				FState(R, (C + R) Mod FNb) = Temp(R, C)
			Next
		Next
	End Sub
	
	Private Sub MixColumns()
		Dim Temp(3, 3)
		Dim R, C
		For R = 0 To 3
			For C = 0 To 3
				Temp(R, C) = FState(R, C)
			Next
		Next
		For C = 0 To 3
			FState(0, C) = CByte(CInt(gfmultby02(Temp(0, C))) Xor CInt(gfmultby03(Temp(1, C))) Xor CInt(gfmultby01(Temp(2, C))) Xor CInt(gfmultby01(Temp(3, C))))
			FState(1, C) = CByte(CInt(gfmultby01(Temp(0, C))) Xor CInt(gfmultby02(Temp(1, C))) Xor CInt(gfmultby03(Temp(2, C))) Xor CInt(gfmultby01(Temp(3, C))))
			FState(2, C) = CByte(CInt(gfmultby01(Temp(0, C))) Xor CInt(gfmultby01(Temp(1, C))) Xor CInt(gfmultby02(Temp(2, C))) Xor CInt(gfmultby03(Temp(3, C))))
			FState(3, C) = CByte(CInt(gfmultby03(Temp(0, C))) Xor CInt(gfmultby01(Temp(1, C))) Xor CInt(gfmultby01(Temp(2, C))) Xor CInt(gfmultby02(Temp(3, C))))
		Next
	End Sub
	
	Private Sub InvMixColumns()
		Dim Temp(3, 3)
		Dim R, C
		For R = 0 To 3
			For C = 0 To 3
				Temp(R, C) = FState(R, C)
			Next
		Next
		For C = 0 To 3
			FState(0, C) = CByte(CInt(gfmultby0e(Temp(0, C))) Xor CInt(gfmultby0b(Temp(1, C))) Xor CInt(gfmultby0d(Temp(2, C))) Xor CInt(gfmultby09(Temp(3, C))))
			FState(1, C) = CByte(CInt(gfmultby09(Temp(0, C))) Xor CInt(gfmultby0e(Temp(1, C))) Xor CInt(gfmultby0b(Temp(2, C))) Xor CInt(gfmultby0d(Temp(3, C))))
			FState(2, C) = CByte(CInt(gfmultby0d(Temp(0, C))) Xor CInt(gfmultby09(Temp(1, C))) Xor CInt(gfmultby0e(Temp(2, C))) Xor CInt(gfmultby0b(Temp(3, C))))
			FState(3, C) = CByte(CInt(gfmultby0b(Temp(0, C))) Xor CInt(gfmultby0d(Temp(1, C))) Xor CInt(gfmultby09(Temp(2, C))) Xor CInt(gfmultby0e(Temp(3, C))))
		Next
	End Sub
	
	Private Function gfmultby01(B)
		gfmultby01 = B
	End Function
	
	Private Function gfmultby02(B)
		If (B < &H80) then gfmultby02 = CByte(CInt(B * 2)) _
		Else gfmultby02 = CByte((CInt(B * 2)) Xor (CInt(&H1b)))
	End Function
	
	Private Function gfmultby03(B)
		gfmultby03 = CByte((CInt(gfmultby02(B))) Xor (CInt(B)))
	End Function
	
	Private Function gfmultby09(B)
		gfmultby09 = CByte((CInt(gfmultby02(gfmultby02(gfmultby02(B))))) Xor (CInt(B)))
	End Function
	
	Private Function gfmultby0b(B)
		gfmultby0b = CByte((CInt(gfmultby02(gfmultby02(gfmultby02(B))))) Xor (CInt(gfmultby02(B))) Xor (CInt(B)))
	End Function
	
	Private Function gfmultby0d(B)
		gfmultby0d = CByte((CInt(gfmultby02(gfmultby02(gfmultby02(B))))) Xor (CInt(gfmultby02(gfmultby02(B)))) Xor (CInt(B)))
	End Function
	
	Private Function gfmultby0e(B)
		gfmultby0e = CByte((CInt(gfmultby02(gfmultby02(gfmultby02(B))))) Xor (CInt(gfmultby02(gfmultby02(B)))) Xor (CInt(gfmultby02(B))))
	End Function
	
	Private Sub SubWord(B1, B2, B3, B4)
		B4 = FSbox(B4 \ 16, B4 And &H0f )
		B3 = FSbox(B3 \ 16, B3 And &H0f )
		B2 = FSbox(B2 \ 16, B2 And &H0f )
		B1 = FSbox(B1 \ 16, B1 And &H0f )
	End Sub
	
	Private Sub RotWord(B1, B2, B3, B4)
		Dim B: B = B1: B1 = B2: B2 = B3: B3 = B4: B4 = B
	End Sub
	
	Private Sub FCipher(Input, Output)
		Dim I, around
		For I = 0 To 4 * FNb - 1
			FState(I Mod 4, I \ 4) = Input(I)
		Next
		AddRoundKey 0
		For around = 1 To FNr - 1
			SubBytes()
			ShIftRows()
			MixColumns()
			AddRoundKey around
		Next
		SubBytes()
		ShIftRows()
		AddRoundKey FNr
		For I = 0 To FNb * 4 - 1
			Output(I) = FState(I Mod 4, I \ 4)
		Next
	End Sub
	
	Private Sub FInvCipher(Input, Output)
		Dim I, around
		For I = 0 To 4 * FNb - 1
			FState(I Mod 4, I \ 4) = Input(I)
		Next
		AddRoundKey FNr
		around = FNr - 1
		Do While around >= 1
			InvShIftRows()
			InvSubBytes()
			AddRoundKey around
			InvMixColumns()
			around = around -1
		Loop
		InvShIftRows()
		InvSubBytes()
		AddRoundKey 0	
		For I = 0 To FNb * 4 - 1
			Output(I) = FState(I Mod 4, I \ 4)
		Next
	End Sub
	
	Private Function BuildSBox()
		FSBox(00, 00) = &H63: FSBox(00, 01) = &H7C: FSBox(00, 02) = &H77: FSBox(00, 03) = &H7B: FSBox(00, 04) = &HF2: FSBox(00, 05) = &H6B: FSBox(00, 06) = &H6F: FSBox(00, 07) = &HC5: FSBox(00, 08) = &H30: FSBox(00, 09) = &H01: FSBox(00, 10) = &H67: FSBox(00, 11) = &H2B: FSBox(00, 12) = &HFE: FSBox(00, 13) = &HD7: FSBox(00, 14) = &HAB: FSBox(00, 15) = &H76
		FSBox(01, 00) = &HCA: FSBox(01, 01) = &H82: FSBox(01, 02) = &HC9: FSBox(01, 03) = &H7D: FSBox(01, 04) = &HFA: FSBox(01, 05) = &H59: FSBox(01, 06) = &H47: FSBox(01, 07) = &HF0: FSBox(01, 08) = &HAD: FSBox(01, 09) = &HD4: FSBox(01, 10) = &HA2: FSBox(01, 11) = &HAF: FSBox(01, 12) = &H9C: FSBox(01, 13) = &HA4: FSBox(01, 14) = &H72: FSBox(01, 15) = &HC0
		FSBox(02, 00) = &HB7: FSBox(02, 01) = &HFD: FSBox(02, 02) = &H93: FSBox(02, 03) = &H26: FSBox(02, 04) = &H36: FSBox(02, 05) = &H3F: FSBox(02, 06) = &HF7: FSBox(02, 07) = &HCC: FSBox(02, 08) = &H34: FSBox(02, 09) = &HA5: FSBox(02, 10) = &HE5: FSBox(02, 11) = &HF1: FSBox(02, 12) = &H71: FSBox(02, 13) = &HD8: FSBox(02, 14) = &H31: FSBox(02, 15) = &H15
		FSBox(03, 00) = &H04: FSBox(03, 01) = &HC7: FSBox(03, 02) = &H23: FSBox(03, 03) = &HC3: FSBox(03, 04) = &H18: FSBox(03, 05) = &H96: FSBox(03, 06) = &H05: FSBox(03, 07) = &H9A: FSBox(03, 08) = &H07: FSBox(03, 09) = &H12: FSBox(03, 10) = &H80: FSBox(03, 11) = &HE2: FSBox(03, 12) = &HEB: FSBox(03, 13) = &H27: FSBox(03, 14) = &HB2: FSBox(03, 15) = &H75
		FSBox(04, 00) = &H09: FSBox(04, 01) = &H83: FSBox(04, 02) = &H2C: FSBox(04, 03) = &H1A: FSBox(04, 04) = &H1B: FSBox(04, 05) = &H6E: FSBox(04, 06) = &H5A: FSBox(04, 07) = &HA0: FSBox(04, 08) = &H52: FSBox(04, 09) = &H3B: FSBox(04, 10) = &HD6: FSBox(04, 11) = &HB3: FSBox(04, 12) = &H29: FSBox(04, 13) = &HE3: FSBox(04, 14) = &H2F: FSBox(04, 15) = &H84
		FSBox(05, 00) = &H53: FSBox(05, 01) = &HD1: FSBox(05, 02) = &H00: FSBox(05, 03) = &HED: FSBox(05, 04) = &H20: FSBox(05, 05) = &HFC: FSBox(05, 06) = &HB1: FSBox(05, 07) = &H5B: FSBox(05, 08) = &H6A: FSBox(05, 09) = &HCB: FSBox(05, 10) = &HBE: FSBox(05, 11) = &H39: FSBox(05, 12) = &H4A: FSBox(05, 13) = &H4C: FSBox(05, 14) = &H58: FSBox(05, 15) = &HCF
		FSBox(06, 00) = &HD0: FSBox(06, 01) = &HEF: FSBox(06, 02) = &HAA: FSBox(06, 03) = &HFB: FSBox(06, 04) = &H43: FSBox(06, 05) = &H4D: FSBox(06, 06) = &H33: FSBox(06, 07) = &H85: FSBox(06, 08) = &H45: FSBox(06, 09) = &HF9: FSBox(06, 10) = &H02: FSBox(06, 11) = &H7F: FSBox(06, 12) = &H50: FSBox(06, 13) = &H3C: FSBox(06, 14) = &H9F: FSBox(06, 15) = &HA8
		FSBox(07, 00) = &H51: FSBox(07, 01) = &HA3: FSBox(07, 02) = &H40: FSBox(07, 03) = &H8F: FSBox(07, 04) = &H92: FSBox(07, 05) = &H9D: FSBox(07, 06) = &H38: FSBox(07, 07) = &HF5: FSBox(07, 08) = &HBC: FSBox(07, 09) = &HB6: FSBox(07, 10) = &HDA: FSBox(07, 11) = &H21: FSBox(07, 12) = &H10: FSBox(07, 13) = &HFF: FSBox(07, 14) = &HF3: FSBox(07, 15) = &HD2
		FSBox(08, 00) = &HCD: FSBox(08, 01) = &H0C: FSBox(08, 02) = &H13: FSBox(08, 03) = &HEC: FSBox(08, 04) = &H5F: FSBox(08, 05) = &H97: FSBox(08, 06) = &H44: FSBox(08, 07) = &H17: FSBox(08, 08) = &HC4: FSBox(08, 09) = &HA7: FSBox(08, 10) = &H7E: FSBox(08, 11) = &H3D: FSBox(08, 12) = &H64: FSBox(08, 13) = &H5D: FSBox(08, 14) = &H19: FSBox(08, 15) = &H73
		FSBox(09, 00) = &H60: FSBox(09, 01) = &H81: FSBox(09, 02) = &H4F: FSBox(09, 03) = &HDC: FSBox(09, 04) = &H22: FSBox(09, 05) = &H2A: FSBox(09, 06) = &H90: FSBox(09, 07) = &H88: FSBox(09, 08) = &H46: FSBox(09, 09) = &HEE: FSBox(09, 10) = &HB8: FSBox(09, 11) = &H14: FSBox(09, 12) = &HDE: FSBox(09, 13) = &H5E: FSBox(09, 14) = &H0B: FSBox(09, 15) = &HDB
		FSBox(10, 00) = &HE0: FSBox(10, 01) = &H32: FSBox(10, 02) = &H3A: FSBox(10, 03) = &H0A: FSBox(10, 04) = &H49: FSBox(10, 05) = &H06: FSBox(10, 06) = &H24: FSBox(10, 07) = &H5C: FSBox(10, 08) = &HC2: FSBox(10, 09) = &HD3: FSBox(10, 10) = &HAC: FSBox(10, 11) = &H62: FSBox(10, 12) = &H91: FSBox(10, 13) = &H95: FSBox(10, 14) = &HE4: FSBox(10, 15) = &H79
		FSBox(11, 00) = &HE7: FSBox(11, 01) = &HC8: FSBox(11, 02) = &H37: FSBox(11, 03) = &H6D: FSBox(11, 04) = &H8D: FSBox(11, 05) = &HD5: FSBox(11, 06) = &H4E: FSBox(11, 07) = &HA9: FSBox(11, 08) = &H6C: FSBox(11, 09) = &H56: FSBox(11, 10) = &HF4: FSBox(11, 11) = &HEA: FSBox(11, 12) = &H65: FSBox(11, 13) = &H7A: FSBox(11, 14) = &HAE: FSBox(11, 15) = &H08
		FSBox(12, 00) = &HBA: FSBox(12, 01) = &H78: FSBox(12, 02) = &H25: FSBox(12, 03) = &H2E: FSBox(12, 04) = &H1C: FSBox(12, 05) = &HA6: FSBox(12, 06) = &HB4: FSBox(12, 07) = &HC6: FSBox(12, 08) = &HE8: FSBox(12, 09) = &HDD: FSBox(12, 10) = &H74: FSBox(12, 11) = &H1F: FSBox(12, 12) = &H4B: FSBox(12, 13) = &HBD: FSBox(12, 14) = &H8B: FSBox(12, 15) = &H8A
		FSBox(13, 00) = &H70: FSBox(13, 01) = &H3E: FSBox(13, 02) = &HB5: FSBox(13, 03) = &H66: FSBox(13, 04) = &H48: FSBox(13, 05) = &H03: FSBox(13, 06) = &HF6: FSBox(13, 07) = &H0E: FSBox(13, 08) = &H61: FSBox(13, 09) = &H35: FSBox(13, 10) = &H57: FSBox(13, 11) = &HB9: FSBox(13, 12) = &H86: FSBox(13, 13) = &HC1: FSBox(13, 14) = &H1D: FSBox(13, 15) = &H9E
		FSBox(14, 00) = &HE1: FSBox(14, 01) = &HF8: FSBox(14, 02) = &H98: FSBox(14, 03) = &H11: FSBox(14, 04) = &H69: FSBox(14, 05) = &HD9: FSBox(14, 06) = &H8E: FSBox(14, 07) = &H94: FSBox(14, 08) = &H9B: FSBox(14, 09) = &H1E: FSBox(14, 10) = &H87: FSBox(14, 11) = &HE9: FSBox(14, 12) = &HCE: FSBox(14, 13) = &H55: FSBox(14, 14) = &H28: FSBox(14, 15) = &HDF
		FSBox(15, 00) = &H8C: FSBox(15, 01) = &HA1: FSBox(15, 02) = &H89: FSBox(15, 03) = &H0D: FSBox(15, 04) = &HBF: FSBox(15, 05) = &HE6: FSBox(15, 06) = &H42: FSBox(15, 07) = &H68: FSBox(15, 08) = &H41: FSBox(15, 09) = &H99: FSBox(15, 10) = &H2D: FSBox(15, 11) = &H0F: FSBox(15, 12) = &HB0: FSBox(15, 13) = &H54: FSBox(15, 14) = &HBB: FSBox(15, 15) = &H16
	End Function
	
	Private Function BuildIsBox()
		FIsBox(00, 00) = &H52: FIsBox(00, 01) = &H09: FIsBox(00, 02) = &H6A: FIsBox(00, 03) = &HD5: FIsBox(00, 04) = &H30: FIsBox(00, 05) = &H36: FIsBox(00, 06) = &HA5: FIsBox(00, 07) = &H38: FIsBox(00, 08) = &HBF: FIsBox(00, 09) = &H40: FIsBox(00, 10) = &HA3: FIsBox(00, 11) = &H9E: FIsBox(00, 12) = &H81: FIsBox(00, 13) = &HF3: FIsBox(00, 14) = &HD7: FIsBox(00, 15) = &HFB 
		FIsBox(01, 00) = &H7C: FIsBox(01, 01) = &HE3: FIsBox(01, 02) = &H39: FIsBox(01, 03) = &H82: FIsBox(01, 04) = &H9B: FIsBox(01, 05) = &H2F: FIsBox(01, 06) = &HFF: FIsBox(01, 07) = &H87: FIsBox(01, 08) = &H34: FIsBox(01, 09) = &H8E: FIsBox(01, 10) = &H43: FIsBox(01, 11) = &H44: FIsBox(01, 12) = &HC4: FIsBox(01, 13) = &HDE: FIsBox(01, 14) = &HE9: FIsBox(01, 15) = &HCB
		FIsBox(02, 00) = &H54: FIsBox(02, 01) = &H7B: FIsBox(02, 02) = &H94: FIsBox(02, 03) = &H32: FIsBox(02, 04) = &HA6: FIsBox(02, 05) = &HC2: FIsBox(02, 06) = &H23: FIsBox(02, 07) = &H3D: FIsBox(02, 08) = &HEE: FIsBox(02, 09) = &H4C: FIsBox(02, 10) = &H95: FIsBox(02, 11) = &H0B: FIsBox(02, 12) = &H42: FIsBox(02, 13) = &HFA: FIsBox(02, 14) = &HC3: FIsBox(02, 15) = &H4E
		FIsBox(03, 00) = &H08: FIsBox(03, 01) = &H2E: FIsBox(03, 02) = &HA1: FIsBox(03, 03) = &H66: FIsBox(03, 04) = &H28: FIsBox(03, 05) = &HD9: FIsBox(03, 06) = &H24: FIsBox(03, 07) = &HB2: FIsBox(03, 08) = &H76: FIsBox(03, 09) = &H5B: FIsBox(03, 10) = &HA2: FIsBox(03, 11) = &H49: FIsBox(03, 12) = &H6D: FIsBox(03, 13) = &H8B: FIsBox(03, 14) = &HD1: FIsBox(03, 15) = &H25
		FIsBox(04, 00) = &H72: FIsBox(04, 01) = &HF8: FIsBox(04, 02) = &HF6: FIsBox(04, 03) = &H64: FIsBox(04, 04) = &H86: FIsBox(04, 05) = &H68: FIsBox(04, 06) = &H98: FIsBox(04, 07) = &H16: FIsBox(04, 08) = &HD4: FIsBox(04, 09) = &HA4: FIsBox(04, 10) = &H5C: FIsBox(04, 11) = &HCC: FIsBox(04, 12) = &H5D: FIsBox(04, 13) = &H65: FIsBox(04, 14) = &HB6: FIsBox(04, 15) = &H92
		FIsBox(05, 00) = &H6C: FIsBox(05, 01) = &H70: FIsBox(05, 02) = &H48: FIsBox(05, 03) = &H50: FIsBox(05, 04) = &HFD: FIsBox(05, 05) = &HED: FIsBox(05, 06) = &HB9: FIsBox(05, 07) = &HDA: FIsBox(05, 08) = &H5E: FIsBox(05, 09) = &H15: FIsBox(05, 10) = &H46: FIsBox(05, 11) = &H57: FIsBox(05, 12) = &HA7: FIsBox(05, 13) = &H8D: FIsBox(05, 14) = &H9D: FIsBox(05, 15) = &H84
		FIsBox(06, 00) = &H90: FIsBox(06, 01) = &HD8: FIsBox(06, 02) = &HAB: FIsBox(06, 03) = &H00: FIsBox(06, 04) = &H8C: FIsBox(06, 05) = &HBC: FIsBox(06, 06) = &HD3: FIsBox(06, 07) = &H0A: FIsBox(06, 08) = &HF7: FIsBox(06, 09) = &HE4: FIsBox(06, 10) = &H58: FIsBox(06, 11) = &H05: FIsBox(06, 12) = &HB8: FIsBox(06, 13) = &HB3: FIsBox(06, 14) = &H45: FIsBox(06, 15) = &H06
		FIsBox(07, 00) = &HD0: FIsBox(07, 01) = &H2C: FIsBox(07, 02) = &H1E: FIsBox(07, 03) = &H8F: FIsBox(07, 04) = &HCA: FIsBox(07, 05) = &H3F: FIsBox(07, 06) = &H0F: FIsBox(07, 07) = &H02: FIsBox(07, 08) = &HC1: FIsBox(07, 09) = &HAF: FIsBox(07, 10) = &HBD: FIsBox(07, 11) = &H03: FIsBox(07, 12) = &H01: FIsBox(07, 13) = &H13: FIsBox(07, 14) = &H8A: FIsBox(07, 15) = &H6B
		FIsBox(08, 00) = &H3A: FIsBox(08, 01) = &H91: FIsBox(08, 02) = &H11: FIsBox(08, 03) = &H41: FIsBox(08, 04) = &H4F: FIsBox(08, 05) = &H67: FIsBox(08, 06) = &HDC: FIsBox(08, 07) = &HEA: FIsBox(08, 08) = &H97: FIsBox(08, 09) = &HF2: FIsBox(08, 10) = &HCF: FIsBox(08, 11) = &HCE: FIsBox(08, 12) = &HF0: FIsBox(08, 13) = &HB4: FIsBox(08, 14) = &HE6: FIsBox(08, 15) = &H73
		FIsBox(09, 00) = &H96: FIsBox(09, 01) = &HAC: FIsBox(09, 02) = &H74: FIsBox(09, 03) = &H22: FIsBox(09, 04) = &HE7: FIsBox(09, 05) = &HAD: FIsBox(09, 06) = &H35: FIsBox(09, 07) = &H85: FIsBox(09, 08) = &HE2: FIsBox(09, 09) = &HF9: FIsBox(09, 10) = &H37: FIsBox(09, 11) = &HE8: FIsBox(09, 12) = &H1C: FIsBox(09, 13) = &H75: FIsBox(09, 14) = &HDF: FIsBox(09, 15) = &H6E
		FIsBox(10, 00) = &H47: FIsBox(10, 01) = &HF1: FIsBox(10, 02) = &H1A: FIsBox(10, 03) = &H71: FIsBox(10, 04) = &H1D: FIsBox(10, 05) = &H29: FIsBox(10, 06) = &HC5: FIsBox(10, 07) = &H89: FIsBox(10, 08) = &H6F: FIsBox(10, 09) = &HB7: FIsBox(10, 10) = &H62: FIsBox(10, 11) = &H0E: FIsBox(10, 12) = &HAA: FIsBox(10, 13) = &H18: FIsBox(10, 14) = &HBE: FIsBox(10, 15) = &H1B
		FIsBox(11, 00) = &HFC: FIsBox(11, 01) = &H56: FIsBox(11, 02) = &H3E: FIsBox(11, 03) = &H4B: FIsBox(11, 04) = &HC6: FIsBox(11, 05) = &HD2: FIsBox(11, 06) = &H79: FIsBox(11, 07) = &H20: FIsBox(11, 08) = &H9A: FIsBox(11, 09) = &HDB: FIsBox(11, 10) = &HC0: FIsBox(11, 11) = &HFE: FIsBox(11, 12) = &H78: FIsBox(11, 13) = &HCD: FIsBox(11, 14) = &H5A: FIsBox(11, 15) = &HF4
		FIsBox(12, 00) = &H1F: FIsBox(12, 01) = &HDD: FIsBox(12, 02) = &HA8: FIsBox(12, 03) = &H33: FIsBox(12, 04) = &H88: FIsBox(12, 05) = &H07: FIsBox(12, 06) = &HC7: FIsBox(12, 07) = &H31: FIsBox(12, 08) = &HB1: FIsBox(12, 09) = &H12: FIsBox(12, 10) = &H10: FIsBox(12, 11) = &H59: FIsBox(12, 12) = &H27: FIsBox(12, 13) = &H80: FIsBox(12, 14) = &HEC: FIsBox(12, 15) = &H5F
		FIsBox(13, 00) = &H60: FIsBox(13, 01) = &H51: FIsBox(13, 02) = &H7F: FIsBox(13, 03) = &HA9: FIsBox(13, 04) = &H19: FIsBox(13, 05) = &HB5: FIsBox(13, 06) = &H4A: FIsBox(13, 07) = &H0D: FIsBox(13, 08) = &H2D: FIsBox(13, 09) = &HE5: FIsBox(13, 10) = &H7A: FIsBox(13, 11) = &H9F: FIsBox(13, 12) = &H93: FIsBox(13, 13) = &HC9: FIsBox(13, 14) = &H9C: FIsBox(13, 15) = &HEF
		FIsBox(14, 00) = &HA0: FIsBox(14, 01) = &HE0: FIsBox(14, 02) = &H3B: FIsBox(14, 03) = &H4D: FIsBox(14, 04) = &HAE: FIsBox(14, 05) = &H2A: FIsBox(14, 06) = &HF5: FIsBox(14, 07) = &HB0: FIsBox(14, 08) = &HC8: FIsBox(14, 09) = &HEB: FIsBox(14, 10) = &HBB: FIsBox(14, 11) = &H3C: FIsBox(14, 12) = &H83: FIsBox(14, 13) = &H53: FIsBox(14, 14) = &H99: FIsBox(14, 15) = &H61
		FIsBox(15, 00) = &H17: FIsBox(15, 01) = &H2B: FIsBox(15, 02) = &H04: FIsBox(15, 03) = &H7E: FIsBox(15, 04) = &HBA: FIsBox(15, 05) = &H77: FIsBox(15, 06) = &HD6: FIsBox(15, 07) = &H26: FIsBox(15, 08) = &HE1: FIsBox(15, 09) = &H69: FIsBox(15, 10) = &H14: FIsBox(15, 11) = &H63: FIsBox(15, 12) = &H55: FIsBox(15, 13) = &H21: FIsBox(15, 14) = &H0C: FIsBox(15, 15) = &H7D
	End Function
	
	Private Function BuildRcon()
		FRcon(00, 00) = &H00: FRcon(00, 01) = &H00: FRcon(00, 02) = &H00: FRcon(00, 03) = &H00
		FRcon(01, 00) = &H01: FRcon(01, 01) = &H00: FRcon(01, 02) = &H00: FRcon(01, 03) = &H00
		FRcon(02, 00) = &H02: FRcon(02, 01) = &H00: FRcon(02, 02) = &H00: FRcon(02, 03) = &H00
		FRcon(03, 00) = &H04: FRcon(03, 01) = &H00: FRcon(03, 02) = &H00: FRcon(03, 03) = &H00
		FRcon(04, 00) = &H08: FRcon(04, 01) = &H00: FRcon(04, 02) = &H00: FRcon(04, 03) = &H00
		FRcon(05, 00) = &H10: FRcon(05, 01) = &H00: FRcon(05, 02) = &H00: FRcon(05, 03) = &H00
		FRcon(06, 00) = &H20: FRcon(06, 01) = &H00: FRcon(06, 02) = &H00: FRcon(06, 03) = &H00
		FRcon(07, 00) = &H40: FRcon(07, 01) = &H00: FRcon(07, 02) = &H00: FRcon(07, 03) = &H00
		FRcon(08, 00) = &H80: FRcon(08, 01) = &H00: FRcon(08, 02) = &H00: FRcon(08, 03) = &H00
		FRcon(09, 00) = &H1B: FRcon(09, 01) = &H00: FRcon(09, 02) = &H00: FRcon(09, 03) = &H00
		FRcon(10, 00) = &H36: FRcon(10, 01) = &H00: FRcon(10, 02) = &H00: FRcon(10, 03) = &H00
	End Function
End Class
%>