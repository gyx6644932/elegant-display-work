<%
'**************************************************
'��������gotTopic
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str		----ԭ�ַ���
'		 strlen		----��ȡ����
'����ֵ����ȡ����ַ���
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
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl		----��ַ
'����ֵ������ ? �� & ����ַ
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
'��������strLength
'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��  ����str		----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
'**************************************************
Function strLength(str)
	If IsNull(str) Or str = "" Then strLength = 0 : Exit Function
    On Error Resume Next
    Dim WINNT_CHINESE
    WINNT_CHINESE = (Len("�й�") = 2)
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
'��������GB2UTF8
'��  �ã�GB2312תUTF8
'��  ����str		----Ҫת�����ַ���
'����ֵ��UTF8
'**************************************************
Function GB2UTF8(Str)
	For i = 1 to Len (Str)
		GB2UTF8 = GB2UTF8 & "&#x" & Hex(Ascw(Mid(Str, i, 1))) & ";"
	next
End Function
'**************************************************
'��������GetStrLength
'��  �ã�UTF-8���ַ������ȡ������������ַ���Ӣ��������һ���ַ���
'��  ����str		----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
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
'��������IsValidPassword
'��  �ã�����ַ������Ƿ��������ַ�
'��  ����str		----Ҫ�����ַ���
'����ֵ��True		----�������ַ�
'		 False		----�������ַ�
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
'��������HasChinese
'��  �ã�����ַ������Ƿ��������ַ�
'��  ����str		----Ҫ�����ַ���
'����ֵ��True		----������
'		 False		----������
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
'��������ReplaceReg		ȫ���滻
'��������ReplaceTest	ֻ�滻��һ��
'��  �ã������滻
'��  ����str1		----Ҫ�滻���ַ���
'		 patrn		----���滻���ַ�
'		 replStr	----����ַ�
'����ֵ���滻����ַ���
'**************************************************
Function ReplaceReg(str1, patrn, replStr)
 Dim regEx
 Set regEx = New RegExp						'����������ʽ��
 regEx.Pattern = patrn						'ģʽ��
 regEx.IgnoreCase = False					'�Ƿ����ִ�Сд��
 regEx.Global = True						'�Ƿ�ȫ��ƥ�䡣
 ReplaceReg = regEx.Replace(str1, replStr)	'���滻��
End Function
Function ReplaceTest(str1, patrn, replStr)	'ȥ��ȫ��ƥ��ֻ�滻��һ��
 Dim regEx
 Set regEx = New RegExp 
 regEx.Pattern = patrn 
 regEx.IgnoreCase = False
 regEx.Global = False
 ReplaceTest = regEx.Replace(str1, replStr)
End Function
'**************************************************
'��������Html2Txt
'��  �ã���HTML����ת�����ı�������
'��  ����str		----Ҫ�滻�Ĵ���
'����ֵ����TXT�ı����ж���
'**************************************************
Function Html2Txt(str)
    If IsNull(str) Or Trim(str) = "" Then
        Html2Txt = ""
        Exit Function
    End If
	str=ReplaceTest(str,"&nbsp;&nbsp;","����")
	str=ReplaceReg(str,"&nbsp;&nbsp;&nbsp;&nbsp;","����")
	str=ReplaceReg(str,"&nbsp;","")
	str=ReplaceReg(str,"(\<.[^\<]*\>)","")
	str=ReplaceReg(str,"(\<\/[^\<]*\>)","")
	str=ReplaceReg(str,"����","<br>����")
	str=ReplaceTest(str,"<br>","")
	Html2Txt=str
End Function
'**************************************************
'��������Clearnum
'��  �ã�������תΪ�ֺ�
'��  ����str		----Ҫת�����ַ���
'����ֵ�����Ϊ�ֺŵ�����
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
'��������Outtxt
'��  �ã�����html Ԫ��
'��  ����str		---- Ҫ�����ַ�
'����ֵ�����ı�
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
	str=Replace(str,Chr(33), "��")
	str=Replace(str,Chr(34), "��")
	str=Replace(str,Chr(40), "��")
	str=Replace(str,Chr(41), "��")
    Outtxt = str
End Function
'**************************************************
'��������txtnormal
'��  �ã������ı�Ԫ��
'��  ����str		---- Ҫ�����ַ�
'����ֵ����������ı�
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
	str=Replace(str,Chr(33), "��")
	str=Replace(str,Chr(34), "��")
	str=Replace(str,Chr(40), "��")
	str=Replace(str,Chr(41), "��")
	str=Replace(str,"&nbsp;&nbsp;", "����")
    txtnormal = str
End Function
'**************************************************
'��������FoundInArr
'��  �ã����һ������������Ԫ���Ƿ����ָ���ַ���
'��  ����strArr		----�洢�������ݵ��ִ�
'        strToFind	----Ҫ���ҵ��ַ���
'        strSplit	----����ķָ���
'����ֵ��True		----��ָ���ַ���
'		 False		----��ָ���ַ���
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
'��������ReplaceBadChar
'��  �ã����˷Ƿ���SQL�ַ�
'��  ����strChar	-----Ҫ���˵��ַ�
'����ֵ�����˺���ַ�
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
'��������SafeChar
'��  �ã��������
'��  ����strChar	-----Ҫ���˵��ַ�
'����ֵ�����˺���ַ�
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
	tempChar = Replace(tempChar,Chr(33),"��")
	tempChar = Replace(tempChar,Chr(34),"��")
	tempChar = Replace(tempChar,Chr(35),"��")
	tempChar = Replace(tempChar,Chr(36),"��")
	tempChar = Replace(tempChar,Chr(37),"��")
	tempChar = Replace(tempChar,Chr(38),"��")
	tempChar = Replace(tempChar,Chr(39),"��")
	tempChar = Replace(tempChar,Chr(40),"��")
	tempChar = Replace(tempChar,Chr(41),"��")
	tempChar = Replace(tempChar,Chr(42),"��")
	tempChar = Replace(tempChar,Chr(43),"��")
	tempChar = Replace(tempChar,Chr(44),"��")
	tempChar = Replace(tempChar,Chr(45),"��")
	tempChar = Replace(tempChar,Chr(46),"��")
	tempChar = Replace(tempChar,Chr(47),"��")
	tempChar = Replace(tempChar,Chr(58),"��")
	tempChar = Replace(tempChar,Chr(59),"��")
	tempChar = Replace(tempChar,Chr(60),"��")
	tempChar = Replace(tempChar,Chr(61),"��")
	tempChar = Replace(tempChar,Chr(62),"��")
	tempChar = Replace(tempChar,Chr(63),"��")
	tempChar = Replace(tempChar,Chr(64),"��")
	tempChar = Replace(tempChar,Chr(91),"��")
	tempChar = Replace(tempChar,Chr(92),"��")
	tempChar = Replace(tempChar,Chr(93),"��")
	tempChar = Replace(tempChar,Chr(94),"��")
	tempChar = Replace(tempChar,Chr(95),"��")
	tempChar = Replace(tempChar,Chr(96),"��")
	tempChar = Replace(tempChar,Chr(123),"��")
	tempChar = Replace(tempChar,Chr(124),"��")
	tempChar = Replace(tempChar,Chr(125),"��")
	tempChar = Replace(tempChar,Chr(126),"��")
    SafeChar = tempChar
End Function
'**************************************************
'��������isPasPic
'��  �ã�����ǿ���ж�
'��  ����str		----����
'����ֵ��1��2��3	----1�ͣ�2�У�3��
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
'��������isTel
'��  �ã���֤�绰�����棩��ʽ�����Ҵ���(2��3λ)-����(2��3λ)-�绰����(7��8λ)-�ֻ���(3λ)" 
'��  ����str		----�绰����
'����ֵ��True		----�Ϸ�
'		 False		----���Ϸ�
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
'��������isMob
'��  �ã���֤�ֻ������ʽ�����Ҵ���(2��3λ)-����ͷ3λ+(8λ)" 
'��  ����str		----�ֻ�����
'����ֵ��True		----�Ϸ�
'		 False		----���Ϸ�
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
'��������IsValidEmail
'��  �ã����Email��ַ�Ϸ���
'��  ����email		----Ҫ����Email��ַ
'����ֵ��True		----Email��ַ�Ϸ�
'		 False		----Email��ַ���Ϸ�
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
'��������nothtml
'��  �ã�����html Ԫ��
'��  ����str		---- Ҫ�����ַ�
'����ֵ��û��html������ַ���
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
'��������ReplaceBadUrl
'��  �ã����˷Ƿ�Url��ַ����
'��  ����strContent	---- Ҫ����Url��ַ
'����ֵ����ȫ��Url��ַ
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
'��������FormatTime
'��  �ã���ʽ��ʱ��
'��  ����TestTime	---- Ҫ��ʽ����ʱ��
'		 style		---- ���ʱ���ʽ
'			1		---- ��ɫ yyyy��mm��dd��hhʱ
'			2		---- dd��00:00:00
'			3		---- yyyy��mm��dd��
'			4		---- yyyy/mm/dd
'			5		---- yyyy-mm-dd hh:mm
'			6		---- yyyy��mm��dd�� hh:mm
'			7		---- yyyymmddhhmmss
'			8		---- yyyy-mm-dd
'			9		---- yyyymmdd
'			0		---- mm-dd
'			10		---- yyyy/mm/dd hh:mm:ss(JSʱ�����)
'����ֵ��ת�����ʱ���ʽ
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
		FormatTime = n&"��"&y&"��"&r&"��"&s&"ʱ"
	Elseif style = 2 Then
		FormatTime = n&"-"&y&"-"&r
	Elseif style = 3 Then
		FormatTime = n&"��"&y&"��"&r&"��"
	Elseif style = 4 Then
		FormatTime = n&"/"&y&"/"&r
	Elseif style = 5 then
		FormatTime = n&"-"&y&"-"&r&" "&s&":"&f
	Elseif style = 6 then
		FormatTime = n&"��"&y&"��"&r&"��"&s&":"& f
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
'��������SearchFieldValue
'��  �ã���һ�������ж��û������һ���ֶε�ֵ�Ƿ��Ѵ���
'��  ����vTableName ---- ����
'		 vFieldName ---- �ֶ�
'		 vFieldValue---- ֵ
'����ֵ��True		---- ����
'		 False		---- ������
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
'��������SearchFieldValue
'��  �ã���һ�������ж��û������һ���ֶε�ֵ�Ƿ��Ѵ��ڣ���������������
'��  ����vTableName	---- ����
'		 vFieldName ---- �ֶ�
'		 vFieldValue---- ֵ
'		 vIDName	---- ��ֵͬID
'		 intIDValue ---- ����ID
'����ֵ��True		---- ����
'		 False		---- ������
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
'��������checkIDCard
'��  �ã�������֤����Ϸ���
'��  ����idcard		---- Ҫ�������֤����
'����ֵ��-1			---- ��ȷ���֤
'**************************************************
Function checkIDCard(idcard) '-1Ϊ��ȷ�����֤������Ϊ�Ƿ����֤ 
Dim Y, JYM 
Dim S, M 
Dim area 
area = "11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,91" 
Dim ereg 
Set ereg = New regexp 
'�������� 
If InStr(1, area, Mid(idcard, 1, 2)) = 0 Then checkIDCard = 1: Exit Function 
'��ݺ���λ������ʽ���� 
Select Case Len(idcard) 
Case 15 
If ((CInt(Mid(idcard, 7, 2)) + 1900) Mod 4 = 0 or ((CInt(Mid(idcard, 7, 2)) + 1900) Mod 100 = 0 And (CInt(Mid(idcard, 7, 2)) + 1900) Mod 4 = 0)) Then 
ereg.Pattern = "^[1-9][0-9]{5}[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))[0-9]{3}$" ';//���Գ������ڵĺϷ��� 
Else 
ereg.Pattern = "^[1-9][0-9]{5}[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|1[0-9]|2[0-8]))[0-9]{3}$" ';//���Գ������ڵĺϷ��� 
End If 
If (ereg.test(idcard)) Then 
checkIDCard = -1 
Else 
checkIDCard = 2 
End If 
Case 18 
'//18λ��ݺ����� 
'//�������ڵĺϷ��Լ�� 
If ((CInt(Mid(idcard, 7, 2)) + 1900) Mod 4 = 0 or ((CInt(Mid(idcard, 7, 2)) + 1900) Mod 100 = 0 And (CInt(Mid(idcard, 7, 2)) + 1900) Mod 4 = 0)) Then 
ereg.Pattern = "^[1-9][0-9]{5}19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))[0-9]{3}[0-9Xx]$" ';//����������ڵĺϷ���������ʽ 
Else 
ereg.Pattern = "^[1-9][0-9]{5}19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|1[0-9]|2[0-8]))[0-9]{3}[0-9Xx]$" ';//ƽ��������ڵĺϷ���������ʽ 
End If 
If (ereg.test(idcard)) Then 
'//����У��λ 
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
'��������GetIp
'��  �ã��õ�IP
'����ֵ��IP��ַ
'**************************************************
function GetIp()
	dim getclientip
	'����ͻ������˴������������Ӧ����ServerVariables("HTTP_X_FORWARDED_FOR")����
	getclientip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If getclientip = "" Then
		'����ͻ���û�ô���Ӧ����Request.ServerVariables("REMOTE_ADDR")����
		getclientip = Request.ServerVariables("REMOTE_ADDR")
	end if
	GetIp = getclientip
end function
'**************************************************
'��������getFileName
'��  �ã�ȡ���ļ�����
'��  ����pgn	---- ·��(Ϊ��ʱ��ȡ��ǰ�ļ���ҳ��)
'����ֵ���ļ�����
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
'��������strj_no
'��  �ã���ʱ��ȡ�ò��ظ��ı���
'��  ����str		---- ��ǰʱ��
'����ֵ�����ظ��ı���
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
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True		----�Ѿ���װ
'        False		----û�а�װ
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
'��������CheckDir
'��  �ã����ĳһ���Ŀ¼�Ƿ����
'��  ����FolderPath ----Ŀ¼��
'��  ����oType		----1:����·��  2�����·��
'����ֵ��True		----����
'        False		----������
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
'��������MakeNewsDir
'��  �ã�����ָ����������Ŀ¼
'��  ����foldername ----Ŀ¼��
'**************************************************
Function MakeNewsDir(foldername)
	dim HowFSO,f
	Set HowFSO = Server.CreateObject("Scripting.FileSystemObject")
    Set f = HowFSO.CreateFolder(foldername)
    MakeNewsDir = True
	Set HowFSO = nothing
End Function
'**************************************************
'��������AutoJump
'��  �ã�ҳ���Զ���ת
'��  ����str1		----	��ʾ��ʾ����
'        url		----	����ת����
'**************************************************
Sub AutoJump(str1, url)
    Response.Write("<br/>&nbsp;&nbsp;<font color=red>"&str1&"</font><br/>")
    Response.Write("<br/>&nbsp;&nbsp;������ת...<br/>")
    Response.Write("<br/>&nbsp;&nbsp;ҳ��û���Զ���ת<a href="&url&">�������</a><br/>")
    Response.Write("<meta http-equiv=refresh content=2;url='"&url&"'>")
End Sub
'**************************************************
'��������PageControl(enPageControl)
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����iCount			----	�ܼ�¼����
'        pagecount		----	�ɷ�ҳ��
'        page			----	ÿҳ��¼����
'        iPageSize		----	��ǰҳ��
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
    '������һҳ��һҳ����
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
	Response.Write("��ǰ��<b><font color=#ff0000>" & page & "</font></b>/<b>" & pageCount & "</b>ҳ"& vbCrLf)
    Response.Write(" ÿҳ<b>" & iPageSize & "</b>��" & vbCrLf)
	Response.Write(" ��<b>" & iCount & "</b>��" & vbCrLf)
	Response.Write("</TD><TD align='right' style='border:none;padding-right:10px;' style='text-align:right;'>" & vbCrLf )
    If Page<= 1 Then
    Else
        Response.Write("<A HREF="&action&"?"&temp&"Page=1 title='��ҳ'><font style='font-weight:bold;'>&lt;&lt;</font></A> " & vbCrLf)
    End If
	if xpage * 10 > 0 then Response.Write("<A HREF=" & action & "?" & temp & "Page="&Cstr(xpage * 10)&" title='��ʮҳ'><font style='font-weight:bold;'>&lt;</font></a> ")
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
	if xx < pageCount then Response.Write("<A HREF="&action&"?"&temp&"Page="&Cstr(xx)&" title='��ʮҳ'><font style='font-weight:bold;'>&gt;</font></a> ")
	if Page>=pageCount then
	else
		Response.Write("<A HREF="&action&"?"&temp&"Page="&pagecount&" title='βҳ'><font style='font-weight:bold;'>&gt;&gt;</font></a> ")
	end if
    Response.Write(" ת��" & "<INPUT style='font-size:12px;border:1px solid #CCCCCC;padding:0px 2px;height:18px;width:30px;line-height:16px;' TYEP='TEXT' NAME='page' SIZE='3' VALUE="&page&">"&"ҳ"&vbCrLf&"<INPUT style='font-size:12px;height:18px;line-height:16px;color:#FFFFFF;PADDING:0px 2px;background-color:#1B1C1D;border-left:#FFFFFF 1px solid;border-top:#FFFFFF 1px solid;border-right:#999999 1px solid;border-bottom:#999999 1px solid;' type='submit' value='ת��'>")
    Response.Write("</TD>" & vbCrLf )
    Response.Write("</TR></form>" & vbCrLf )
    Response.Write("</table>" & vbCrLf )
End Sub
'**************************************************
'��������strfroms
'��  �ã�����ȫ�ֱ���ֵ
'��  ����str	----	��ֵname,�á�||��������
'		 str	----	��ֵvalue
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
'��������GetTotalSize
'��  �ã��鿴��С
'��  ����GetLocal	----����(��·��)
'        GetType	----����
'        			"File"		----�ļ�
'        			"Folder"	----�ļ���
'����ֵ����С(**KB �� **MB)
'**************************************************
Function GetTotalSize(GetLocal,GetType)
	Set HowFSO=Server.CreateObject("Scripting.FileSystemObject")
	If Err<>0 Then
		Err.Clear
		GetTotalSize="��������֧��FSO����ȡ�ļ���Сʧ�ܣ�"
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
'��������CheckFileExists
'��  �ã��ж��ļ��Ƿ����(֧������·��������·�������·��)
'��  ����chkFiles	----�ļ�����(��·��)
'����ֵ��True		----����
'        False		----������
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
'��������CopyFiles
'��  �ã������ļ�
'��  ����TempSource	----Դ�ļ�����(��·��)
'		 TempEnd	----Ŀ���ļ�����(��·��)
'**************************************************
Function CopyFiles(TempSource,TempEnd)
	IF CheckFileExists(TempEnd) THEN
		response.write "<SCRIPT LANGUAGE=JavaScript>alert ('Ŀ���ļ���" & TempEnd & "���Ѵ��ڣ�����ɾ����');history.back(-1);</script>"
		response.end
		Exit Function
	End IF
	If Not CheckFileExists(TempSource) Then
		response.write "<SCRIPT LANGUAGE=JavaScript>alert ('Դ�ļ���" & TempSource & "�������ڣ�');history.back(-1);</script>"
		response.end
		Exit Function
	End If
	Dim HowFSO
	Set HowFSO = Server.CreateObject("Scripting.FileSystemObject")
	HowFSO.CopyFile TempSource,TempEnd
	Set HowFSO = Nothing
End Function
'**************************************************
'��������DeleteFiles
'��  �ã�ɾ���ļ�
'��  ����TempSource	----�ļ�����(��·��)
'**************************************************
Function DeleteFiles(TempSource)
	If Not CheckFileExists(TempSource) Then
		response.write "<SCRIPT LANGUAGE=JavaScript>alert ('Դ�ļ���" & TempSource & "�������ڣ�');history.back(-1);</script>"
		response.end
		Exit Function
	End If
	Dim HowFSO
	Set HowFSO = Server.CreateObject("Scripting.FileSystemObject")
	HowFSO.DeleteFile TempSource
	Set HowFSO = Nothing
End Function
'**************************************************
'��������FSOreName
'��  �ã��������ļ�
'��  ����sourceName	----Դ�ļ�(��·��)
'		 destName	----���ļ���
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
'��������gotseat
'��  �ã��Զ���ȡ���
'��  ����insum		----����
'		 inpeo		----��ȡ����ţ�������ʱ����Ž���������������ſ�2��Ϊ�ѣ�
'		 inseat		----���ŵı�ţ���ʽ��*,001,002,003,*,004,005��088,089,*,091,*,099,100,������ȡ���ſ�2��
'����ֵ��ʣ��ı��|ռ���ı��
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
'������NewOrder() 
'������sz	����
'	   op	�ָ���
'	   nu	��λ
'����������������������� 
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
'��������gotTonum
'��  �ã��������֣����飩��д
'��  ����str		----ԭ�������֣����飩
'		 op			----�ָ���
'����ֵ�������д����������֣����飩
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
					t = t&" �� "&c
					if i < thei then
						t = t&op&Clng(trim(strTemp(i)))
					end if
				end if
				if i = thei then
					if Clng(trim(strTemp(thei))) - Clng(trim(strTemp(thei - 1))) = 1 then
						t = t&" �� "&Clng(trim(strTemp(i)))
					else
						t = t&op&Clng(trim(strTemp(i)))
					end if
				end if
			next
			strTemp = split(t,op)
			t = ""
			for i = 0 to Ubound(strTemp)
				if Ubound(split(strTemp(i),"��")) = 1 then
					if trim(split(strTemp(i),"��")(0)) = trim(split(strTemp(i),"��")(1)) then
						t = t&trim(split(strTemp(i),"��")(0))
					else
						t = t&trim(split(strTemp(i),"��")(0))&" �� "&trim(split(strTemp(i),"��")(1))
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
'��������isshuzi
'��  �ã��Ƿ�Ϊ����
'��  ����str		----ԭ�ַ���
'����ֵ�����֣�true,�����֣�false
'**************************************************
Function isshuzi(str)
	isshuzi = false
    If trim(str) <> "" and isNumeric(str) Then isshuzi = true
End Function
'**************************************************
'��������jsouterr
'��  �ã�js�������
'��  ����str		----������������
'����ֵ�����Ѵ��󲢺���һҳ
'**************************************************
Function jsouterr(str)
	Response.write("<script type='text/javascript'>alert('"&str&"'); history.go(-1);</script>")
	Response.End()
End Function
'**************************************************
'��������jsoutgo
'��  �ã�js������ҳ
'��  ����str1		----������������
'		 str2		----��ҳ��ַ
'����ֵ�����Ѵ����ҳ
'**************************************************
Function jsoutgo(str1,str2)
	Response.write("<script type='text/javascript'>alert('"&str1&"'); location.href='"&str2&"';</script>")
    Response.End
End Function
'**************************************************
'��������ShowMain
'��  �ã�ȡ���ݿ������
'��  ����main	----����
'		 code	----�ֶ�
'		 where	----����

'����ֵ����ȡ����ַ���
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
'��������go_Num
'��  �ã��Ƿ�Ϊ���֣���������str2����,����ṩ��str2Ҳ�����֣������0
'��  ����str1		----ԭ�ַ���
'		 str2		----Ҫ�滻���ַ���
'����ֵ�����֣�str1,�����֣�str2
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
'��������ssform
'��  �ã�ͬ��name�ύ��ʽ��
'��  ����str1		----�����ύ����name
'����ֵ����ʽ�������ݣ��á�,���ָ�
'**************************************************
Function ssform(str)
	dim x,y
	y = ""
	for each x in Request.Form(str)
		if go_num(x,0) <> 0 then
			x = go_num(x,0)
		else
			x = replace(x,",","��")
		end if
		y = y&x&","
	next
	if strLength(y) > 0 then y = left(y,len(y) - 1)
	ssform = y
End Function
%>