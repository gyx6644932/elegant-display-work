<%
'**************************************************
'函数名：gotTopic
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str		----原字符串
'		 strlen		----截取长度
'返回值：截取后的字符串
'**************************************************
Function gotTopic(ByVal str, ByVal strlen)
    If trim(str) <> "" Then
		Dim l, t, c, i, strTemp
		str = Replace(Replace(Replace(Replace(str, "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
		l = Len(str)
		t = 0
		strTemp = str
		strlen = CLng(strlen)
		For i = 1 To l
			c = Abs(Asc(Mid(str, i, 1)))
			If c > 255 Then
				t = t + 2
			Else
				t = t + 1
			End If
			If t >= strlen Then
				strTemp = Left(str, i)
				Exit For
			End If
		Next
		If strTemp <> str Then
			strTemp = strTemp & "..."
		End If
		gotTopic = Replace(Replace(Replace(Replace(strTemp, " ", "&nbsp;"), Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
	Else
	    gotTopic = ""
        Exit Function
    End If
End Function
'**************************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl		----网址
'返回值：加了 ? 或 & 的网址
'**************************************************
Function JoinChar(ByVal strUrl)
    If strUrl = "" Then
        JoinChar = ""
        Exit Function
    End If
    If InStr(strUrl, "?") < Len(strUrl) Then
        If InStr(strUrl, "?") > 1 Then
            If InStr(strUrl, "&") < Len(strUrl) Then
                JoinChar = strUrl & "&"
            Else
                JoinChar = strUrl
            End If
        Else
            JoinChar = strUrl & "?"
        End If
    Else
        JoinChar = strUrl
    End If
End Function
'**************************************************
'函数名：strLength
'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
'参  数：str		----要求长度的字符串
'返回值：字符串长度
'**************************************************
Function strLength(str)
	If IsNull(str) Or str = "" Then strLength = 0 : Exit Function
    On Error Resume Next
    Dim WINNT_CHINESE
    WINNT_CHINESE = (Len("中国") = 2)
    If WINNT_CHINESE Then
        Dim l, t, c
        Dim i
        l = Len(str)
        t = l
        For i = 1 To l
            c = Asc(Mid(str, i, 1))
            If c < 0 Then c = c + 65536
            If c > 255 Then
                t = t + 1
            End If
        Next
        strLength = t
    Else
        strLength = Len(str)
    End If
    If Err.Number <> 0 Then Err.Clear
End Function
'**************************************************
'函数名：GB2UTF8
'作  用：GB2312转UTF8
'参  数：str		----要转换的字符串
'返回值：UTF8
'**************************************************
Function GB2UTF8(Str)
	For i = 1 to Len (Str)
		GB2UTF8 = GB2UTF8 & "&#x" & Hex(Ascw(Mid(Str, i, 1))) & ";"
	next
End Function
'**************************************************
'函数名：GetStrLength
'作  用：UTF-8求字符串长度。汉字算两个字符，英文数字算一个字符。
'参  数：str		----要求长度的字符串
'返回值：字符串长度
'**************************************************
Public Function GetStrLength(ByVal Str)
Dim oRegExp,TemStr
If IsNull(Str) Or Str = "" Then GetStrLength = 0 : Exit Function
Set oRegExp = New RegExp
oRegExp.IgnoreCase = True
oRegExp.Global = True
oRegExp.Pattern = "[\uff00-\uffff\u4e00-\u9fa5\ufe10-\ufe1f\ufe30-\ufe4f\u1100-\u11ff\u2600-\u26ff\u2700-\u27bf\u2800-\u28ff\u3300-\u33ff\u3200-\u32ff\ua490-\ua4cf\ua000-\ua48f\u3130-\u318f\uac00-\ud7af\u31f0-\u31ff\u30a0-\u30ff\u3040-\u309f\u31a0-\u31bf\u3100-\u312F\u2FF0-\u2FFF\u2F00-\u2FDF\u31c0-\u31ef\u3000-\u303f\u2e80-\u2eff\uff00-\uffef]"
TemStr = oRegExp.Replace(str, "xx")
Set oRegExp = Nothing
GetStrLength = Len(TemStr)
End Function
'**************************************************
'函数名：IsValidPassword
'作  用：检查字符串中是否有特殊字符
'参  数：str		----要检查的字符串
'返回值：True		----有特殊字符
'		 False		----无特殊字符
'**************************************************
Public Function IsValidPassword(ByVal str)
	IsValidPassword = True
	On Error Resume Next
	If IsNull(str) Then Exit Function
	If Trim(str) = Empty Then Exit Function
	Dim ForbidStr, i
	ForbidStr = "=|&|<|>|?|%|,|;|:|(|)|`|~|!|*|#|$|^|{|}|[|]|+|-|/|\|" & Chr(32) & "|" & Chr(34) & "|" & Chr(39) & "|" & Chr(9)
	ForbidStr = Split(ForbidStr, "|")
	For i = 0 To UBound(ForbidStr)
		If InStr(1, str, ForbidStr(i), 1) > 0 Then
			IsValidPassword = True
			Exit Function
		End If
	Next
	IsValidPassword = False
End Function
'**************************************************
'函数名：HasChinese
'作  用：检查字符串中是否有中文字符
'参  数：str		----要检查的字符串
'返回值：True		----有中文
'		 False		----无中文
'**************************************************
function HasChinese(str) 
HasChinese = false 
dim i 
for i=1 to Len(str) 
if Asc(Mid(str,i,1)) < 0 then 
HasChinese = true 
exit for 
end if 
next 
end function
'**************************************************
'函数名：ReplaceReg		全程替换
'函数名：ReplaceTest	只替换第一个
'作  用：正则替换
'参  数：str1		----要替换的字符串
'		 patrn		----欲替换的字符
'		 replStr	----结果字符
'返回值：替换后的字符串
'**************************************************
Function ReplaceReg(str1, patrn, replStr)
 Dim regEx
 Set regEx = New RegExp						'建立正则表达式。
 regEx.Pattern = patrn						'模式。
 regEx.IgnoreCase = False					'是否区分大小写。
 regEx.Global = True						'是否全程匹配。
 ReplaceReg = regEx.Replace(str1, replStr)	'作替换。
End Function
Function ReplaceTest(str1, patrn, replStr)	'去掉全程匹配只替换第一个
 Dim regEx
 Set regEx = New RegExp 
 regEx.Pattern = patrn 
 regEx.IgnoreCase = False
 regEx.Global = False
 ReplaceTest = regEx.Replace(str1, replStr)
End Function
'**************************************************
'函数名：Html2Txt
'作  用：将HTML代码转换成文本段落型
'参  数：str		----要替换的代码
'返回值：纯TXT文本带有段落
'**************************************************
Function Html2Txt(str)
    If IsNull(str) Or Trim(str) = "" Then
        Html2Txt = ""
        Exit Function
    End If
	str=ReplaceTest(str,"&nbsp;&nbsp;","　　")
	str=ReplaceReg(str,"&nbsp;&nbsp;&nbsp;&nbsp;","　　")
	str=ReplaceReg(str,"&nbsp;","")
	str=ReplaceReg(str,"(\<.[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/[^\<]*\>)","")
	str=ReplaceReg(str,"　　","<br>　　")
	str=ReplaceTest(str,"<br>","")
	Html2Txt=str
End Function
'**************************************************
'函数名：Clearnum
'作  用：数组间隔转为分号
'参  数：str		----要转换的字符串
'返回值：间隔为分号的数组
'**************************************************
function Clearnum(str)
	str=replace(str,",",";")
	str=replace(str," ",";")
	for i=0 to Ubound(Split(str,";"))+1
		str=replace(str,";;",";")
		str=replace(str,";;",";")
		if not isNumeric(right(str,1)) then str=left(str,len(str)-1)
		if not isNumeric(left(str,1)) then str=right(str,len(str)-1)
	next
	Clearnum=str
end function
'**************************************************
'函数名：Outtxt
'作  用：过滤html 元素
'参  数：str		---- 要过滤字符
'返回值：纯文本
'**************************************************
Function Outtxt(str)
    If IsNull(str) Or Trim(str) = "" Then
        Outtxt = ""
        Exit Function
    End If
	str=ReplaceReg(str,"&nbsp;","")
	str=ReplaceReg(str,"(\<.[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/[^\<]*\>)","")
	str=ReplaceReg(str,"	","")
	str=Replace(str,"'", "")
	str=Replace(str,Chr(10), "")
	str=Replace(str,Chr(13), "")
	str=Replace(str,Chr(33), "！")
	str=Replace(str,Chr(34), "“")
	str=Replace(str,Chr(40), "（")
	str=Replace(str,Chr(41), "）")
    Outtxt = str
End Function
'**************************************************
'函数名：txtnormal
'作  用：过滤文本元素
'参  数：str		---- 要过滤字符
'返回值：美化后的文本
'**************************************************
Function txtnormal(str)
    If IsNull(str) Or Trim(str) = "" Then
        txtnormal = ""
        Exit Function
    End If
	str=ReplaceReg(str,"(\<font[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/font[^\<]*\>)","")
	str=ReplaceReg(str,"(\<FONT[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/FONT[^\<]*\>)","")
	str=ReplaceReg(str,"(\<span[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/span[^\<]*\>)","")
	str=ReplaceReg(str,"(\<SPAN[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/SPAN[^\<]*\>)","")
	str=ReplaceReg(str,"(\<div[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/div[^\<]*\>)","")
	str=ReplaceReg(str,"(\<DIV[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/DIV[^\<]*\>)","")
	str=ReplaceReg(str,"(\<o[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/o[^\<]*\>)","")
	str=ReplaceReg(str,"(\<O[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/O[^\<]*\>)","")
	str=ReplaceReg(str,"(\<b [^\<]*\>)","<b>")
	str=ReplaceReg(str,"(\<B [^\<]*\>)","<b>")
	str=ReplaceReg(str,"(\<st1[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/st1[^\<]*\>)","")
	str=ReplaceReg(str,"(\<p[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/p[^\<]*\>)","<br>")
	str=ReplaceReg(str,"(\<P[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/P[^\<]*\>)","<br>")
	str=ReplaceReg(str,"(\<\?xml[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\?XML[^\<]*\>)","")
	str=ReplaceReg(str,"(\<a [^\<]*\>)","")
	str=ReplaceReg(str,"(\<A [^\<]*\>)","")
	str=Replace(str,"'", "")
	str=Replace(str,Chr(10), "")
	str=Replace(str,Chr(13), "")
	str=Replace(str,Chr(33), "！")
	str=Replace(str,Chr(34), "“")
	str=Replace(str,Chr(40), "（")
	str=Replace(str,Chr(41), "）")
	str=Replace(str,"&nbsp;&nbsp;", "　　")
    txtnormal = str
End Function
'**************************************************
'函数名：FoundInArr
'作  用：检查一个数组中所有元素是否包含指定字符串
'参  数：strArr		----存储数据数据的字串
'        strToFind	----要查找的字符串
'        strSplit	----数组的分隔符
'返回值：True		----有指定字符串
'		 False		----无指定字符串
'**************************************************
Function FoundInArr(strArr, strToFind, strSplit)
    Dim arrTemp, i
    FoundInArr = False
    If InStr(strArr, strSplit) > 0 Then
        arrTemp = Split(strArr, strSplit)
        For i = 0 To UBound(arrTemp)
        If LCase(Trim(arrTemp(i))) = LCase(Trim(strToFind)) Then
            FoundInArr = True
            Exit For
        End If
        Next
    Else
        If LCase(Trim(strArr)) = LCase(Trim(strToFind)) Then
        FoundInArr = True
        End If
    End If
End Function
'**************************************************
'函数名：ReplaceBadChar
'作  用：过滤非法的SQL字符
'参  数：strChar	-----要过滤的字符
'返回值：过滤后的字符
'**************************************************
Function ReplaceBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',--,%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceBadChar = tempChar
End Function
'**************************************************
'函数名：SafeChar
'作  用：保存过虑
'参  数：strChar	-----要过滤的字符
'返回值：过滤后的字符
'**************************************************
Function SafeChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        SafeChar = ""
        Exit Function
    End If
    Dim tempChar
	tempChar = strChar
	tempChar = Replace(tempChar,Chr(9)," ")
	tempChar = Replace(tempChar,Chr(10)," ")
	tempChar = Replace(tempChar,Chr(13)," ")
	tempChar = Replace(tempChar,Chr(32)," ")
	tempChar = Replace(tempChar,Chr(33),"！")
	tempChar = Replace(tempChar,Chr(34),"“")
	tempChar = Replace(tempChar,Chr(35),"＃")
	tempChar = Replace(tempChar,Chr(36),"＄")
	tempChar = Replace(tempChar,Chr(37),"％")
	tempChar = Replace(tempChar,Chr(38),"＆")
	tempChar = Replace(tempChar,Chr(39),"‘")
	tempChar = Replace(tempChar,Chr(40),"（")
	tempChar = Replace(tempChar,Chr(41),"）")
	tempChar = Replace(tempChar,Chr(42),"※")
	tempChar = Replace(tempChar,Chr(43),"＋")
	tempChar = Replace(tempChar,Chr(44),"，")
	tempChar = Replace(tempChar,Chr(45),"－")
	tempChar = Replace(tempChar,Chr(46),"．")
	tempChar = Replace(tempChar,Chr(47),"／")
	tempChar = Replace(tempChar,Chr(58),"：")
	tempChar = Replace(tempChar,Chr(59),"；")
	tempChar = Replace(tempChar,Chr(60),"＜")
	tempChar = Replace(tempChar,Chr(61),"＝")
	tempChar = Replace(tempChar,Chr(62),"＞")
	tempChar = Replace(tempChar,Chr(63),"？")
	tempChar = Replace(tempChar,Chr(64),"＠")
	tempChar = Replace(tempChar,Chr(91),"【")
	tempChar = Replace(tempChar,Chr(92),"＼")
	tempChar = Replace(tempChar,Chr(93),"】")
	tempChar = Replace(tempChar,Chr(94),"＾")
	tempChar = Replace(tempChar,Chr(95),"＿")
	tempChar = Replace(tempChar,Chr(96),"｀")
	tempChar = Replace(tempChar,Chr(123),"≮")
	tempChar = Replace(tempChar,Chr(124),"｜")
	tempChar = Replace(tempChar,Chr(125),"≯")
	tempChar = Replace(tempChar,Chr(126),"～")
    SafeChar = tempChar
End Function
'**************************************************
'函数名：isPasPic
'作  用：密码强度判断
'参  数：str		----密码
'返回值：1，2，3	----1低，2中，3高
'**************************************************
Function isPasPic(str)
	isPasPic = 1
	temp = str
	If temp = "" Or IsNull(temp) Then
        Exit Function
	elseif len(temp) < 6 then
        Exit Function
    End If
	dim mos, strtmp, i, temp
	mos = 0
	dim regEx
	Set regEx = New RegExp
	regEx.Pattern = "[0-9]"
	regEx.IgnoreCase = False
	regEx.Global = True
	If regEx.Test(temp) = True Then mos = mos + 1
	temp = regEx.Replace(temp,"")
	regEx.Pattern = "[a-z]"
	If regEx.Test(temp) = True Then mos = mos + 1
	temp = regEx.Replace(temp,"")
	regEx.Pattern = "[A-Z]"
	If regEx.Test(temp) = True Then mos = mos + 1
	temp = regEx.Replace(temp,"")
	if temp <> "" then mos = mos + 1
	if mos > 3 then
		isPasPic = 3
	else
		isPasPic = mos
	end if
	Set regEx = Nothing
End Function
'**************************************************
'函数名：isTel
'作  用：验证电话（传真）格式：国家代码(2到3位)-区号(2到3位)-电话号码(7到8位)-分机号(3位)" 
'参  数：str		----电话号码
'返回值：True		----合法
'		 False		----不合法
'**************************************************
Function isTel(ByVal str)
	Set regEx = New RegExp
	regEx.Pattern = "^(([0\+]\d{2,3}-)?(0\d{2,3})-)?(\d{7,8})(-(\d{3,}))?$"
	regEx.IgnoreCase = True
	regEx.Global = True
	isTel = regEx.Test(str)
	Set regEx = Nothing
End Function
'**************************************************
'函数名：isMob
'作  用：验证手机号码格式：国家代码(2到3位)-号码头3位+(8位)" 
'参  数：str		----手机号码
'返回值：True		----合法
'		 False		----不合法
'**************************************************
Function isMob(ByVal str)
	Set regEx = New RegExp
	regEx.Pattern = "^(([0\+]\d{2,3}-)?(1[358][0-9])+\d{8})$"
	regEx.IgnoreCase = True
	regEx.Global = True
	isMob = regEx.Test(str)
	Set regEx = Nothing
End Function
'**************************************************
'函数名：IsValidEmail
'作  用：检查Email地址合法性
'参  数：email		----要检查的Email地址
'返回值：True		----Email地址合法
'		 False		----Email地址不合法
'**************************************************
Function IsValidEmail(email)
    Dim names, name, i, c
    IsValidEmail = True
    names = Split(email, "@")
    If UBound(names) <> 1 Then
       IsValidEmail = False
       Exit Function
    End If
    For Each name In names
        If Len(name) <= 0 Then
        IsValidEmail = False
        Exit Function
        End If
        For i = 1 To Len(name)
        c = LCase(Mid(name, i, 1))
        If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
           IsValidEmail = False
           Exit Function
         End If
       Next
       If Left(name, 1) = "." Or Right(name, 1) = "." Then
          IsValidEmail = False
          Exit Function
       End If
    Next
    If InStr(names(1), ".") <= 0 Then
        IsValidEmail = False
       Exit Function
    End If
    i = Len(names(1)) - InStrRev(names(1), ".")
    If i <> 2 And i <> 3 And i <> 4 Then
       IsValidEmail = False
       Exit Function
    End If
    If InStr(email, "..") > 0 Then
       IsValidEmail = False
    End If
End Function
'**************************************************
'函数名：nothtml
'作  用：过滤html 元素
'参  数：str		---- 要过滤字符
'返回值：没有html代码的字符串
'**************************************************
Public Function nothtml(ByVal str)
    If IsNull(str) Or Trim(str) = "" Then
        nothtml = ""
        Exit Function
    End If
    Dim re
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "(\<.[^\<]*\>)"
    str = re.Replace(str, " ")
    re.Pattern = "(\<\/[^\<]*\>)"
    str = re.Replace(str, " ")
    Set re = Nothing
    str = Replace(str, "'", "")
    str = Replace(str, Chr(34), "")
    nothtml = str
End Function
'**************************************************
'函数名：ReplaceBadUrl
'作  用：过滤非法Url地址函数
'参  数：strContent	---- 要过滤Url地址
'返回值：安全的Url地址
'**************************************************
Public Function ReplaceBadUrl(ByVal strContent)
    Dim regEx, Matches, Match
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = True
    regEx.Pattern = "(a|%61|%41)(d|%64|%44)(m|%6D|4D)(i|%69|%49)(n|%6E|%4E)(\_|%5F)(.*?)(.|%2E)(a|%61|%41)(s|%73|%53)(p|%70|%50)"
    Set Matches = regEx.Execute(strContent)
    For Each Match In Matches
        strContent = Replace(strContent, Match.Value, "")
    Next
    regEx.Pattern = "(u|%75|%55)(s|%73|%53)(e|%65|%45)(r|%72|%52)(\_|%5F)(.*?)(.|%2E)(a|%61|%41)(s|%73|%53)(p|%70|%50)"
    Set Matches = regEx.Execute(strContent)
    For Each Match In Matches
        strContent = Replace(strContent, Match.Value, "")
    Next
    Set regEx = Nothing
    ReplaceBadUrl = strContent
End Function
'**************************************************
'函数名：FormatTime
'作  用：格式化时间
'参  数：TestTime	---- 要格式化的时间
'		 style		---- 结果时间格式
'			1		---- 红色 yyyy年mm月dd日hh时
'			2		---- dd日00:00:00
'			3		---- yyyy年mm月dd日
'			4		---- yyyy/mm/dd
'			5		---- yyyy-mm-dd hh:mm
'			6		---- yyyy年mm月dd日 hh:mm
'			7		---- yyyymmddhhmmss
'			8		---- yyyy-mm-dd
'			9		---- yyyymmdd
'			0		---- mm-dd
'			10		---- yyyy/mm/dd hh:mm:ss(JS时间对象)
'返回值：转换后的时间格式
'**************************************************
Function FormatTime(TestTime,style)
	Dim n,y,r,s,f,m
	n = Year(TestTime)
	y = Month(TestTime)
	r = Day(TestTime)
	s = Hour(TestTime)
	f = Minute(TestTime)
	m = Second(TestTime)
	if len(n) = 2 then n = "20" & n
	if len(y) = 1 then y = "0" & y
	if len(r) = 1 then r = "0" & r						
	if len(s) = 1 then s = "0" & s
	if len(f) = 1 then f = "0" & f
	if len(m) = 1 then m = "0" & m
	If style = 1 Then
		FormatTime = n&"年"&y&"月"&r&"日"&s&"时"
	Elseif style = 2 Then
		FormatTime = n&"-"&y&"-"&r
	Elseif style = 3 Then
		FormatTime = n&"年"&y&"月"&r&"日"
	Elseif style = 4 Then
		FormatTime = n&"/"&y&"/"&r
	Elseif style = 5 then
		FormatTime = n&"-"&y&"-"&r&" "&s&":"&f
	Elseif style = 6 then
		FormatTime = n&"年"&y&"月"&r&"日"&s&":"& f
	Elseif style = 7 then
		FormatTime = n&y&r&s&f&m
	Elseif style = 8 then
		FormatTime = n&"-"&y&"-"&r
	Elseif style = 9 then
		FormatTime = n&y&r
	Elseif style = 0 then
		FormatTime = y&"-"&r
	Elseif style = 10 then
		FormatTime = n&"/"&y&"/"&r&" "&s&":"&f&":"&m
	End if
End Function
'**************************************************
'函数名：SearchFieldValue
'作  用：在一个表中判断用户输入的一个字段的值是否已存在
'参  数：vTableName ---- 表名
'		 vFieldName ---- 字段
'		 vFieldValue---- 值
'返回值：True		---- 存在
'		 False		---- 不存在
'**************************************************
Function SearchFieldValue(vTableName,vFieldName,vFieldValue)
	dim strField,sqlField
	set strField = Server.CreateObject("ADODB.Recordset")
	sqlField = "Select * From "& vTableName &" Where "& vFieldName &" = '"& vFieldValue &"'"
	strField.Open sqlField,Conn
	if not strField.EOF then
		SearchFieldValue = True
	else
		SearchFieldValue = False
	end if
	strField.Close
	set strField = nothing
end Function
'**************************************************
'函数名：SearchFieldValue
'作  用：在一个表中判断用户输入的一个字段的值是否已存在，除了他本身以外
'参  数：vTableName	---- 表名
'		 vFieldName ---- 字段
'		 vFieldValue---- 值
'		 vIDName	---- 相同值ID
'		 intIDValue ---- 本身ID
'返回值：True		---- 存在
'		 False		---- 不存在
'**************************************************
Function SearchEditFieldValue(vTableName,vFieldname,vFieldValue,vIDName,intIDValue)
	dim strField1,sqlField1
	set strField1 = Server.CreateObject("ADODB.Recordset")
	sqlField1 = "Select * From "& vTableName &" Where "& vFieldName &" = '"& vFieldValue &"'"
	strField1.Open sqlField1,Conn
	if not strField1.EOF then
		do while not strField1.EOF
			if int(intIDValue) <> strField1(vIDName) then
				SearchEditFieldValue = True
				exit Function
			end if
			strField1.MoveNext
		loop
		SearchEditFieldValue = False
	end if
	strField1.Close
	set strField1 = nothing
End Function
'**************************************************
'函数名：checkIDCard
'作  用：检查身份证号码合法性
'参  数：idcard		---- 要检查的身份证号码
'返回值：-1			---- 正确身份证
'**************************************************
Function checkIDCard(idcard) '-1为正确的身份证，否则为非法身份证 
Dim Y, JYM 
Dim S, M 
Dim area 
area = "11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,91" 
Dim ereg 
Set ereg = New regexp 
'地区检验 
If InStr(1, area, Mid(idcard, 1, 2)) = 0 Then checkIDCard = 1: Exit Function 
'身份号码位数及格式检验 
Select Case Len(idcard) 
Case 15 
If ((CInt(Mid(idcard, 7, 2)) + 1900) Mod 4 = 0 or ((CInt(Mid(idcard, 7, 2)) + 1900) Mod 100 = 0 And (CInt(Mid(idcard, 7, 2)) + 1900) Mod 4 = 0)) Then 
ereg.Pattern = "^[1-9][0-9]{5}[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))[0-9]{3}$" ';//测试出生日期的合法性 
Else 
ereg.Pattern = "^[1-9][0-9]{5}[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|1[0-9]|2[0-8]))[0-9]{3}$" ';//测试出生日期的合法性 
End If 
If (ereg.test(idcard)) Then 
checkIDCard = -1 
Else 
checkIDCard = 2 
End If 
Case 18 
'//18位身份号码检测 
'//出生日期的合法性检查 
If ((CInt(Mid(idcard, 7, 2)) + 1900) Mod 4 = 0 or ((CInt(Mid(idcard, 7, 2)) + 1900) Mod 100 = 0 And (CInt(Mid(idcard, 7, 2)) + 1900) Mod 4 = 0)) Then 
ereg.Pattern = "^[1-9][0-9]{5}19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))[0-9]{3}[0-9Xx]$" ';//闰年出生日期的合法性正则表达式 
Else 
ereg.Pattern = "^[1-9][0-9]{5}19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|1[0-9]|2[0-8]))[0-9]{3}[0-9Xx]$" ';//平年出生日期的合法性正则表达式 
End If 
If (ereg.test(idcard)) Then 
'//计算校验位 
S = (CInt(Mid(idcard, 0 + 1, 1)) + CInt(Mid(idcard, 10 + 1, 1))) * 7 _ 
+ (CInt(Mid(idcard, 1 + 1, 1)) + CInt(Mid(idcard, 11 + 1, 1))) * 9 _ 
+ (CInt(Mid(idcard, 2 + 1, 1)) + CInt(Mid(idcard, 12 + 1, 1))) * 10 _ 
+ (CInt(Mid(idcard, 3 + 1, 1)) + CInt(Mid(idcard, 13 + 1, 1))) * 5 _ 
+ (CInt(Mid(idcard, 4 + 1, 1)) + CInt(Mid(idcard, 14 + 1, 1))) * 8 _ 
+ (CInt(Mid(idcard, 5 + 1, 1)) + CInt(Mid(idcard, 15 + 1, 1))) * 4 _ 
+ (CInt(Mid(idcard, 6 + 1, 1)) + CInt(Mid(idcard, 16 + 1, 1))) * 2 _ 
+ CInt(Mid(idcard, 7 + 1, 1)) * 1 _ 
+ CInt(Mid(idcard, 8 + 1, 1)) * 6 _ 
+ CInt(Mid(idcard, 9 + 1, 1)) * 3 
Y = S Mod 11 
M = "F" 
JYM = "10X98765432" 
M = Mid(JYM, Y + 1, 1) 
If (M = Mid(idcard, 17 + 1, 1)) Then checkIDCard = -1 Else checkIDCard = 3 
Else 
checkIDCard = 4 
End If 
Case Else 
checkIDCard = Len(idcard) 
End Select 
End Function
'**************************************************
'函数名：GetIp
'作  用：得到IP
'返回值：IP地址
'**************************************************
function GetIp()
	dim getclientip
	'如果客户端用了代理服务器，则应该用ServerVariables("HTTP_X_FORWARDED_FOR")方法
	getclientip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If getclientip = "" Then
		'如果客户端没用代理，应该用Request.ServerVariables("REMOTE_ADDR")方法
		getclientip = Request.ServerVariables("REMOTE_ADDR")
	end if
	GetIp = getclientip
end function
'**************************************************
'函数名：getFileName
'作  用：取得文件名称
'参  数：pgn	---- 路径(为空时获取当前文件、页名)
'返回值：文件名称
'**************************************************
Function getFileName(pgn)
	pgn = replace(pgn,"\","/")
	if trim(pgn) = "this" then
		apgn = Request.Servervariables("url")
		getFileName = Right(apgn,len(apgn) - InstrRev(apgn,"/"))
	else
		getFileName = Right(pgn,len(pgn) - InstrRev(pgn,"/"))
	end if
End Function
'**************************************************
'函数名：strj_no
'作  用：按时间取得不重复的编码
'参  数：str		---- 当前时间
'返回值：不重复的编码
'**************************************************
function strj_no(str)
	strj_no = ""
	strj_no = strj_no&year(str)
	strj_no = strj_no&right("00"&month(str),2)
	strj_no = strj_no&right("00"&day(str),2)
	strj_no = strj_no&right("00"&hour(str),2)
	strj_no = strj_no&right("00"&minute(str),2)
	strj_no = strj_no&right("00"&second(str),2)
	dim strj_no_ms
	strj_no_ms = timer()*100
	strj_no_ms = right(strj_no_ms,2)
	strj_no_ms = Cint(strj_no_ms) * 100 + 0.5
	strj_no_ms = int(strj_no_ms) / 100
	strj_no_ms = left(strj_no_ms&"00",2)
	strj_no = strj_no&strj_no_ms
end function
'**************************************************
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True		----已经安装
'        False		----没有安装
'**************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
'**************************************************
'函数名：CheckDir
'作  用：检查某一相对目录是否存在
'参  数：FolderPath ----目录名
'参  数：oType		----1:物理路径  2：相对路径
'返回值：True		----存在
'        False		----不存在
'**************************************************
Function CheckDir(FolderPath,oType)
	dim HowFSO
	if oType=1 then
		folderpath = replace(folderpath,"\","/")
	elseif oType=2 then
		folderpath = Server.MapPath(".")&"\"&folderpath
	end if
	Set HowFSO = Server.CreateObject("Scripting.FileSystemObject")
	If HowFSO.FolderExists(FolderPath) then
		CheckDir = True
	Else
		CheckDir = False
	End if
	Set HowFSO = nothing
End Function
'**************************************************
'函数名：MakeNewsDir
'作  用：根据指定名称生成目录
'参  数：foldername ----目录名
'**************************************************
Function MakeNewsDir(foldername)
	dim HowFSO,f
	Set HowFSO = Server.CreateObject("Scripting.FileSystemObject")
    Set f = HowFSO.CreateFolder(foldername)
    MakeNewsDir = True
	Set HowFSO = nothing
End Function
'**************************************************
'过程名：AutoJump
'作  用：页面自动跳转
'参  数：str1		----	显示提示文字
'        url		----	待跳转链接
'**************************************************
Sub AutoJump(str1, url)
    Response.Write("<br/>&nbsp;&nbsp;<font color=red>"&str1&"</font><br/>")
    Response.Write("<br/>&nbsp;&nbsp;正在跳转...<br/>")
    Response.Write("<br/>&nbsp;&nbsp;页面没有自动跳转<a href="&url&">【点这里】</a><br/>")
    Response.Write("<meta http-equiv=refresh content=2;url='"&url&"'>")
End Sub
'**************************************************
'过程名：PageControl(enPageControl)
'作  用：显示“上一页 下一页”等信息
'参  数：iCount			----	总记录条数
'        pagecount		----	可分页数
'        page			----	每页记录条数
'        iPageSize		----	当前页号
'**************************************************
Sub PageControl(iCount, pagecount, page, iPageSize)
	%>
	<style type="text/css">
	<!--
	.pagex{border:none;width:100%;margin:0 auto;}
	.pagex a{color:#666666;text-decoration:none;padding:2px 4px 0px 4px;border:1px solid #666666;background-color:#ffffff;}
	.pagex a:hover{color:#FF0000;text-decoration:none;background:#FFCECF;border:1px solid #FF4246;}
	-->
	</style>
	<%
    '生成上一页下一页链接
    Dim query, a, x, temp, action ,xpage ,xx
	xpage = (Page - 1) \ 10
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
    query = Split(Request.ServerVariables("QUERY_STRING"), "&")
    For Each x In query
		if instr(x,"=") > 0 then
			a = Split(x, "=")
			If StrComp(a(0), "page", vbTextCompare) <> 0 Then
				temp = temp & a(0) & "=" & a(1) & "&"
			End If
		end if
    Next
    Response.Write("<table cellpadding='0' cellspacing='0' class='pagex'>" & vbCrLf )
    Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR >" & vbCrLf )
    Response.Write("<TD align='left' style='border:none;padding-left:10px;' style='text-align:left;'>" & vbCrLf )
	Response.Write("当前：<b><font color=#ff0000>" & page & "</font></b>/<b>" & pageCount & "</b>页"& vbCrLf)
    Response.Write(" 每页<b>" & iPageSize & "</b>条" & vbCrLf)
	Response.Write(" 共<b>" & iCount & "</b>条" & vbCrLf)
	Response.Write("</TD><TD align='right' style='border:none;padding-right:10px;' style='text-align:right;'>" & vbCrLf )
    If Page<= 1 Then
    Else
        Response.Write("<A HREF="&action&"?"&temp&"Page=1 title='首页'><font style='font-weight:bold;'>&lt;&lt;</font></A> " & vbCrLf)
    End If
	if xpage * 10 > 0 then Response.Write("<A HREF=" & action & "?" & temp & "Page="&Cstr(xpage * 10)&" title='上十页'><font style='font-weight:bold;'>&lt;</font></a> ")
	Response.Write "<b>"
	for xx = xpage * 10 + 1 to xpage * 10 + 10
		if xx = Page then
			Response.Write("<font color=""#FF0000"">"+Cstr(xx)+"</font> ")
		else
			Response.Write("<a HREF=" & action & "?" & temp & "Page="&Cstr(xx)&">"+Cstr(xx)+"</a> ")
		end if
		if xx = pageCount then exit for
	next
	Response.Write "</b>"
	if xx < pageCount then Response.Write("<A HREF="&action&"?"&temp&"Page="&Cstr(xx)&" title='下十页'><font style='font-weight:bold;'>&gt;</font></a> ")
	if Page>=pageCount then
	else
		Response.Write("<A HREF="&action&"?"&temp&"Page="&pagecount&" title='尾页'><font style='font-weight:bold;'>&gt;&gt;</font></a> ")
	end if
    Response.Write(" 转第" & "<INPUT style='font-size:12px;border:1px solid #CCCCCC;padding:0px 2px;height:18px;width:30px;line-height:16px;' TYEP='TEXT' NAME='page' SIZE='3' VALUE="&page&">"&"页"&vbCrLf&"<INPUT style='font-size:12px;height:18px;line-height:16px;color:#FFFFFF;PADDING:0px 2px;background-color:#1B1C1D;border-left:#FFFFFF 1px solid;border-top:#FFFFFF 1px solid;border-right:#999999 1px solid;border-bottom:#999999 1px solid;' type='submit' value='转到'>")
    Response.Write("</TD>" & vbCrLf )
    Response.Write("</TR></form>" & vbCrLf )
    Response.Write("</table>" & vbCrLf )
End Sub
'**************************************************
'过程名：strfroms
'作  用：保留全局被传值
'参  数：str	----	新值name,用“||”定义多个
'		 str	----	新值value
'**************************************************
Sub strfroms(str, str2)
    Dim sfquery, sfa, sfx, sftemp, sstr1, sstr2, i, okk
    sfquery = Split(Request.ServerVariables("QUERY_STRING"), "&")
	sstr1 = Split(str, "||")
	sstr2 = Split(str2, "||")
	okk = 0
    For Each sfx In sfquery
        sfa = Split(sfx, "=")
		for i = 0 to ubound(sstr1)
			If StrComp(sfa(0), trim(sstr1(i)), vbTextCompare) <> 0 Then
				okk = okk + 0
			else
				okk = okk + 1
			End If
		next
		if okk = 0 then sftemp = sftemp & sfa(0) & "=" & sfa(1) & "&"
    Next
	Response.Write(sftemp)
	for i = 0 to ubound(sstr1)
		Response.Write(trim(sstr1(i)) & "=" & sstr2(i))
		if i < ubound(sstr1) then Response.Write("&")
	next
End Sub
'**************************************************
'过程名：GetTotalSize
'作  用：查看大小
'参  数：GetLocal	----名称(带路径)
'        GetType	----类型
'        			"File"		----文件
'        			"Folder"	----文件夹
'返回值：大小(**KB 或 **MB)
'**************************************************
Function GetTotalSize(GetLocal,GetType)
	Set HowFSO=Server.CreateObject("Scripting.FileSystemObject")
	If Err<>0 Then
		Err.Clear
		GetTotalSize="服务器不支持FSO，获取文件大小失败！"
	Else
		Dim SiteFolder
		If GetType="Folder" Then
			Set SiteFolder=HowFSO.GetFolder(GetLocal)
		Else
			Set SiteFolder=HowFSO.GetFile(GetLocal)
		End If
		GetTotalSize=SiteFolder.Size
		If GetTotalSize>1024*1024Then
			GetTotalSize=GetTotalSize/1024/1024
			If inStr(GetTotalSize,".") Then GetTotalSize =Left(GetTotalSize,inStr(GetTotalSize,".")+2)
				GetTotalSize=GetTotalSize&" MB"
			Else
				GetTotalSize=Fix(GetTotalSize/1024)&" KB"
			End If
			Set SiteFolder=Nothing
		End If
	Set HowFSO=Nothing
End Function
'**************************************************
'过程名：CheckFileExists
'作  用：判断文件是否存在(支持物理路径、虚拟路径、相对路径)
'参  数：chkFiles	----文件名称(带路径)
'返回值：True		----存在
'        False		----不存在
'**************************************************
Function CheckFileExists(chkFiles)
	Dim lpFiles
	CheckFileExists = False
	lpFiles = chkFiles
	If lpFiles = "" Then Exit Function End If
	lpFiles = Replace(LCase(lpFiles),"\","/")
	If InStr(lpFiles,":/") < 1 And InStr(lpFiles,"http://") < 1 Then
	   lpFiles = Server.MapPath(lpFiles)
	End If
	dim HowFSO
	Set HowFSO = Server.CreateObject("Scripting.FileSystemObject")
	If HowFSO.FileExists(lpFiles) Then
	   CheckFileExists = True
	Else
	   CheckFileExists = False
	End If
	Set HowFSO = Nothing
End Function
'**************************************************
'过程名：CopyFiles
'作  用：复制文件
'参  数：TempSource	----源文件名称(带路径)
'		 TempEnd	----目标文件名称(带路径)
'**************************************************
Function CopyFiles(TempSource,TempEnd)
	IF CheckFileExists(TempEnd) THEN
		response.write "<SCRIPT LANGUAGE=JavaScript>alert ('目标文件【" & TempEnd & "】已存在，请先删除！');history.back(-1);</script>"
		response.end
		Exit Function
	End IF
	If Not CheckFileExists(TempSource) Then
		response.write "<SCRIPT LANGUAGE=JavaScript>alert ('源文件【" & TempSource & "】不存在！');history.back(-1);</script>"
		response.end
		Exit Function
	End If
	Dim HowFSO
	Set HowFSO = Server.CreateObject("Scripting.FileSystemObject")
	HowFSO.CopyFile TempSource,TempEnd
	Set HowFSO = Nothing
End Function
'**************************************************
'过程名：DeleteFiles
'作  用：删除文件
'参  数：TempSource	----文件名称(带路径)
'**************************************************
Function DeleteFiles(TempSource)
	If Not CheckFileExists(TempSource) Then
		response.write "<SCRIPT LANGUAGE=JavaScript>alert ('源文件【" & TempSource & "】不存在！');history.back(-1);</script>"
		response.end
		Exit Function
	End If
	Dim HowFSO
	Set HowFSO = Server.CreateObject("Scripting.FileSystemObject")
	HowFSO.DeleteFile TempSource
	Set HowFSO = Nothing
End Function
'**************************************************
'过程名：FSOreName
'作  用：重命名文件
'参  数：sourceName	----源文件(带路径)
'		 destName	----新文件名
'**************************************************
Function FSOreName(sourceName,destName)
	Dim lpFiles,HowFSO,HowFile
	lpFiles = sourceName
	If lpFiles = "" Then Exit Function End If
	lpFiles = Replace(LCase(lpFiles),"\","/")
	Set HowFSO = Server.CreateObject("Scripting.FileSystemObject")
	If InStr(lpFiles,":/") < 1 And InStr(lpFiles,"http://") < 1 Then
	   lpFiles = Server.MapPath(lpFiles)
	End If
	set HowFile=HowFSO.getFile(lpFiles)
	HowFile.Name=destName
	Set HowFSO=Nothing 
	Set HowFile=Nothing 
End Function
'**************************************************
'函数名：gotseat
'作  用：自动获取编号
'参  数：insum		----人数
'		 inpeo		----可取最大编号（无连号时，编号将超过最大名额，故因放宽到2倍为佳）
'		 inseat		----空着的编号（格式：*,001,002,003,*,004,005…088,089,*,091,*,099,100,）库中取，放宽到2倍
'返回值：剩余的编号|占掉的编号
'**************************************************
Function gotseat(insum, inpeo, inseat)
	dim out_seat,seata,seatb,i,ii,ccstr,Seatnumber
	seata = 1
	seatb = insum
	for ii = 0 to inpeo
		Seatnumber = ""
		for i = seata to seatb
			Seatnumber = Seatnumber&right("00"&i,3)&","
		next
		ccstr = "*,"&Seatnumber&"*"
		if Instr(inseat,ccstr) > 0 then
			out_seat = Replace(Replace(inseat,Seatnumber,"*,"),"*,*","*")
			exit for
		else
			if Clng(seatb) - Clng(inpeo) >= 0 then exit for
			seata = seata + 1
			seatb = seatb + 1
		end if
	next
	if seatb - inpeo >= 0 then
		seata = 1
		seatb = insum
	end if
	for ii = 0 to inpeo
		Seatnumber = ""
		for i = seata to seatb
			Seatnumber = Seatnumber&right("00"&i,3)&","
		next
		if Instr(inseat,Seatnumber) > 0 then
			out_seat = Replace(Replace(inseat,Seatnumber,"*,"),"*,*","*")
			exit for
		else
			if Clng(seatb) - Clng(inpeo) >= 0 then exit for
			seata = seata + 1
			seatb = seatb + 1
		end if
	next
	gotseat = out_seat&"|"&Seatnumber
End Function
'****************************** 
'函数：NewOrder() 
'参数：sz	数组
'	   op	分隔符
'	   nu	少位
'描述：对数组进行重新排序 
'****************************** 
Function NewOrder(sz,op,nu) 
Dim ali,icount,i,ii,j,itemp
ali=split(sz,op) 
icount=UBound(ali) - Clng(nu)
For i=0 To icount 
For j=icount - 1 To i Step -1 
If j+1 <= UBound(ali) Then 
If int(ali(j)) > int(ali(j+1)) Then 
itemp=ali(j) 
ali(j)=ali(j+1) 
ali(j+1)=itemp 
End If 
End If 
Next 
Next 
For ii=0 to Ubound(ali) 
If ii = Ubound(ali) Then 
NewOrder = NewOrder & ali(ii) 
Else 
NewOrder = NewOrder & ali(ii) & op
End If 
Next 
End Function
'**************************************************
'函数名：gotTonum
'作  用：连续数字（数组）简写
'参  数：str		----原连续数字（数组）
'		 op			----分隔符
'返回值：输出简写后的连续数字（数组）
'**************************************************
Function gotTonum(ByVal str, ByVal op)
    If trim(str) <> "" Then
		Dim t, c, i, strTemp
		strTemp = split(str,op)
		dim thei
		if isNumeric(trim(strTemp(Ubound(strTemp)))) then
			thei = Ubound(strTemp)
		else
			thei = Ubound(strTemp) - 1
		end if
		if thei > 0 then
			t = Clng(trim(strTemp(0)))
			for i = 1 to thei
				if not isNumeric(trim(strTemp(i))) then
					gotTonum = ""
					Exit Function
				end if
				if Clng(trim(strTemp(i))) - Clng(trim(strTemp(i-1))) <> 1 then
					c = Clng(trim(strTemp(i-1)))
					t = t&" ～ "&c
					if i < thei then
						t = t&op&Clng(trim(strTemp(i)))
					end if
				end if
				if i = thei then
					if Clng(trim(strTemp(thei))) - Clng(trim(strTemp(thei - 1))) = 1 then
						t = t&" ～ "&Clng(trim(strTemp(i)))
					else
						t = t&op&Clng(trim(strTemp(i)))
					end if
				end if
			next
			strTemp = split(t,op)
			t = ""
			for i = 0 to Ubound(strTemp)
				if Ubound(split(strTemp(i),"～")) = 1 then
					if trim(split(strTemp(i),"～")(0)) = trim(split(strTemp(i),"～")(1)) then
						t = t&trim(split(strTemp(i),"～")(0))
					else
						t = t&trim(split(strTemp(i),"～")(0))&" ～ "&trim(split(strTemp(i),"～")(1))
					end if
				else
					t = t&strTemp(i)
				end if
				if i < Ubound(strTemp) then t = t&op
			next
			gotTonum = t
		else
			if isNumeric(trim(strTemp(0))) then
				gotTonum = Clng(trim(strTemp(0)))
			else
				gotTonum = ""
			end if
        	Exit Function
		end if
	Else
	    gotTonum = ""
        Exit Function
    End If
End Function
'**************************************************
'函数名：isshuzi
'作  用：是否为数字
'参  数：str		----原字符串
'返回值：数字：true,非数字：false
'**************************************************
Function isshuzi(str)
	isshuzi = false
    If trim(str) <> "" and isNumeric(str) Then isshuzi = true
End Function
'**************************************************
'函数名：jsouterr
'作  用：js输出错误
'参  数：str		----错误提醒内容
'返回值：提醒错误并后退一页
'**************************************************
Function jsouterr(str)
	Response.write("<script type='text/javascript'>alert('"&str&"'); history.go(-1);</script>")
	Response.End()
End Function
'**************************************************
'函数名：jsoutgo
'作  用：js提醒跳页
'参  数：str1		----错误提醒内容
'		 str2		----跳页地址
'返回值：提醒错后跳页
'**************************************************
Function jsoutgo(str1,str2)
	Response.write("<script type='text/javascript'>alert('"&str1&"'); location.href='"&str2&"';</script>")
    Response.End
End Function
'**************************************************
'函数名：ShowMain
'作  用：取数据库表内容
'参  数：main	----表名
'		 code	----字段
'		 where	----条件

'返回值：截取后的字符串
'**************************************************
Function ShowMain(main,code,where)
	Set rs_main = Server.CreateObject("ADODB.Recordset")
	sql = "select top 1 "&code&" from ["&main&"] "&where
	rs_main.open sql, conn, 1, 1
	If Not rs_main.eof Then
		ShowMain = rs_main(0)
	else
		ShowMain = ""
	End If
	rs_main.close
	Set rs_main = nothing
End Function
'**************************************************
'函数名：go_Num
'作  用：是否为数字，非数字用str2代替,如果提供的str2也非数字，则输出0
'参  数：str1		----原字符串
'		 str2		----要替换的字符串
'返回值：数字：str1,非数字：str2
'**************************************************
Function go_Num(str1,str2)
	if str1 <> "" and isNumeric(str1) then
		go_Num = str1
    else
		if str2 <> "" and isNumeric(str2) then
			go_Num = str2
		else
			go_Num = 0
		end if
	end if
End Function
'**************************************************
'函数名：ssform
'作  用：同名name提交格式化
'参  数：str1		----表单中提交来的name
'返回值：格式化的数据，用“,”分隔
'**************************************************
Function ssform(str)
	dim x,y
	y = ""
	for each x in Request.Form(str)
		if go_num(x,0) <> 0 then
			x = go_num(x,0)
		else
			x = replace(x,",","，")
		end if
		y = y&x&","
	next
	if strLength(y) > 0 then y = left(y,len(y) - 1)
	ssform = y
End Function
%>