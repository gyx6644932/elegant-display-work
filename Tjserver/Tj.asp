<!--#include file="../inc/conn.asp"-->
<!--#include file="ip.asp"-->
<%
response.Charset="GB2312"
Tj_tem = Request.ServerVariables("HTTP_USER_AGENT")
'获得IP地址
ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
if ip = "" or isnull(ip) then ip = Request.ServerVariables("REMOTE_ADDR")
'开始记录访客
temp_str = ""
Tj_user = Request.Cookies("Tjserver")("username")
'已经访问过网站的，计入点击量
if Tj_user <> "" and not isnull(Tj_user) Then
   temp_str = Tj_user
Else
'记录访客，编排、写入客户名
   sql = "select * from Tj_config"
   set Trs = Server.Createobject("adodb.recordset")
   Trs.open sql,conn,1,3
   If Cstr(Cdate(Trs("view_date"))) = Cstr(Date()) then
	  count_id = Clng(Trs("view_dayid")) + 1
	  Trs("view_dayid") = count_id
   Else
      Trs("view_dayid") = 1 
      count_id = 1
      Trs("view_date") = Date()
   End If
   Trs.update
   Trs.close
   Set Trs = Nothing
   '今天新IP则计入IP访问量
   sql="select * from Tj_online where CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),GetDate(),112) and lastip = '"&ip&"'"
   set Trs = Server.Createobject("adodb.recordset")
   Trs.open sql,conn,1,1
   If Trs.eof then
      sql = "select * from Tj_stat where CONVERT(varchar(100),[view_date],112) = CONVERT(varchar(100),GetDate(),112)"
      set Trsb = Server.Createobject("adodb.recordset")
      Trsb.open sql,conn,1,3
      If Trsb.eof Then
         Trsb.addnew
         Trsb("view_date") = date()
	     Trsb("view_number") = 1
      Else
         if Conn.Execute("Select Count(id) From Tj_online where CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),GetDate(),112) and lastip = '"&ip&"'")(0) <= 0 then Trsb("view_number") = Trsb("view_number") + 1
      End If
	  Trsb.update
      Trsb.close
      Set Trsb = Nothing
   End If
   Trs.close
   Set Trs = nothing
   y = year(date())
   m = month(date())
   If Len(m) = 1 Then m = "0"&m
   d = day(date())
   If Len(d) = 1 Then d = "0"&d
   temp_str = y&m&d&count_id
   Response.Cookies("Tjserver")("username") = temp_str
   Tj_user = temp_str
end if
'记录来路页面
vcome = Request.QueryString("vcome")
'记录本站的被访问页面
vpage = Request.QueryString("vpage")
'记录访客来源地址（省市）
vwhere = address1(ip)
'记录访客使用的网络（如电信、联通等）
vwheref = address2(ip)
'记录访客使用的屏幕分辨率
screeninfo = replace(replace(Request.QueryString("screeninfo1"),"'",""),Chr(34),"")&"×"&replace(replace(Request.QueryString("screeninfo2"),"'",""),Chr(34),"")
'识别访客使用的浏览器
browser = Lcase(Tj_tem)
if InStr(browser,"tencent traveler") > 0 or InStr(browser,"tencenttraveler") > 0 then
	browser = "Tencent Traveler(腾讯TT)"
elseif InStr(browser,"alibrowser") > 0 then
	browser = "Alibrowser(阿云)"
elseif InStr(browser,"maxthon") > 0 then
	browser = "Maxthon(傲游)"
elseif InStr(browser,"metasr") > 0 and InStr(browser,"se") > 0 then
	browser = "Sogou Explorer(搜狗)"
elseif InStr(browser,"360se") > 0 then
	browser = "360se(360)"
elseif InStr(browser,"ruibin") > 0 then
	browser = "Rayying(瑞影)"
elseif InStr(browser,"krbrowser") > 0 then
	browser = "Krbrowser(KR)"
elseif InStr(browser,"sleipnir") > 0 then
	browser = "Sleipnir(神马)"
elseif InStr(browser,"qqbrowser") > 0 then
	browser = "QQbrowser(QQ)"
elseif InStr(browser,"slimbrowser") > 0 then
	browser = "Slimbrowser"
elseif InStr(browser,"saayaa") > 0 then
	browser = "Saayaa(闪游)"
elseif InStr(browser,"lunascape") > 0 then
	browser = "Lunascape"
elseif InStr(browser,"dooble") > 0 then
	browser = "Dooble"
elseif InStr(browser,"2345explorer") > 0 then
	browser = "2345 Explorer(2345)"
elseif InStr(browser,"coolnovo") > 0 then
	browser = "Coolnovo(枫树)"
elseif InStr(browser,"msie 8") > 0 then
	browser = "Internet Explorer 8"
elseif InStr(browser,"msie 7") > 0 then
	browser = "Internet Explorer 7"
elseif InStr(browser,"msie 6") > 0 then
	browser = "Internet Explorer 6"
elseif InStr(browser,"msie 5") > 0 then
	browser = "Internet Explorer 5"
elseif InStr(browser,"msie 4") > 0 then
	browser = "Internet Explorer 4"
elseif InStr(browser,"msie 3") > 0 then
	browser = "Internet Explorer 3"
elseif InStr(browser,"firefox") > 0 then
	browser = "FireFox(火狐)"
elseif InStr(browser,"opera") > 0 then
	browser = "Opera(欧朋)"
elseif InStr(browser,"chrome") > 0 then
	browser = "Google Chrome(谷歌)"
elseif InStr(browser,"safari") > 0 then
	browser = "Safari(苹果)"
elseif InStr(browser,"avant") > 0 then
	browser = "Avant"
else
	browser = "未知"
end if
'识别访客使用的系统
systemer = Lcase(Tj_tem)
if Instr(systemer,"nt 6.1") > 0 then
	systemer = "Windows 7"
elseif Instr(systemer,"nt 6.0") > 0 then
	systemer = "Windows Vista"
elseif Instr(systemer,"nt 5.2") > 0 then
	systemer = "Windows 2003"
elseif Instr(systemer,"nt 5.1") > 0 then
	systemer = "Windows XP"
elseif Instr(systemer,"nt 5") > 0 then
	systemer = "Windows 2000"
elseif Instr(systemer,"nt 4") > 0 then
	systemer = "Windows NT4"
elseif Instr(systemer,"4.9") > 0 then
	systemer = "Windows ME"
elseif Instr(systemer,"98") > 0 then
	systemer = "Windows 98"
elseif Instr(systemer,"95") > 0 then
	systemer = "Windows 95"
elseif instr(systemer,"Mac") > 0 then
	systemer = "Mac"	
elseif instr(systemer,"unix") > 0 then
	systemer = "Unix"
elseif instr(systemer,"linux") > 0 then
	systemer = "Linux"
elseif instr(systemer,"sunos") > 0 then
	systemer = "SunOS"
elseif Instr(systemer,"webzip") > 0 Then
	systemer = "webzip"
elseif Instr(systemer,"flashget") > 0 Then
	systemer = "flashget"
elseif Instr(systemer,"offline") > 0 Then
	systemer = "offline"
elseif Instr(systemer,"Tel") > 0 then
	systemer = "Telport"
elseif instr(systemer,"bsd") > 0 then
	systemer = "BSD"
else
	systemer = "未知"
end if
'识别访客使用的搜索引擎
If InStr(vcome, "http://baidu.com") > 0 or InStr(vcome, ".baidu.com") > 0 Then
vcheck = "百度"
elseif InStr(vcome, "http://google.") > 0 or InStr(vcome, ".google.") > 0 Then
vcheck = "谷歌Google"
elseif InStr(vcome, "http://openfind.com") > 0 or InStr(vcome, ".openfind.com") > 0 Then
vcheck = "网擎"
elseif InStr(vcome, "http://lycos.com") > 0 or InStr(vcome, ".lycos.com") > 0 Then
vcheck = "Lycos"
elseif InStr(vcome, "http://search.tom.com") > 0 or InStr(vcome, ".tom.com") > 0 Then
vcheck = "TOM"
elseif InStr(vcome, "http://zhongsou.com") > 0 or InStr(vcome, ".zhongsou.com") > 0 Then
vcheck = "中搜"
elseif InStr(vcome, "http://bing.com") > 0 or InStr(vcome, ".bing.com") > 0 Then
vcheck = "必应"
elseif InStr(vcome, "http://yisou.com") > 0 or InStr(vcome, ".yisou.com") > 0 Then
vcheck = "一搜"
elseif InStr(vcome, "http://sina.") > 0 or InStr(vcome, ".sina.") > 0 Then
vcheck = "新浪爱问"
elseif InStr(vcome, "http://sohu.com") > 0 or InStr(vcome, ".sohu.com") > 0 Then
vcheck = "搜狐"
elseif InStr(vcome, "http://3721.com") > 0 or InStr(vcome, ".3721.com") > 0 Then
vcheck = "3721"
elseif InStr(vcome, "http://soso.com") > 0 or InStr(vcome, ".soso.com") > 0 Then
vcheck = "腾讯搜搜"
elseif InStr(vcome, "http://youdao.com") > 0 or InStr(vcome, ".youdao.com") > 0 Then
vcheck = "网易有道"
elseif InStr(vcome, "http://sogou.com") > 0 or InStr(vcome, ".sogou.com") > 0 Then
vcheck = "搜狗"
elseif InStr(vcome, "http://search.com") > 0 or InStr(vcome, ".search.com") > 0 Then
vcheck = "MSN.Search"
elseif InStr(vcome, "http://search.aol.com") > 0 or InStr(vcome, ".aol.com") > 0 Then
vcheck = "AOL"
elseif InStr(vcome, "http://alexa.com") > 0 or InStr(vcome, ".alexa.com") > 0 Then
vcheck = "Alexa"
elseif InStr(vcome, "http://114.com") > 0 or InStr(vcome, ".114.com") > 0 Then
vcheck = "114"
elseif InStr(vcome, "http://115.com") > 0 or InStr(vcome, ".115.com") > 0 Then
vcheck = "115"
elseif InStr(vcome, "http://qq.com") > 0 or InStr(vcome, ".qq.com") > 0 Then
vcheck = "腾讯"
else
vcheck = "未知"
end if
'识别访客在搜索引擎中使用的关键字
if InStr(vcome,"baidu.") > 0 or InStr(vcome,"google.") > 0 or InStr(vcome,"soso.") > 0 or InStr(vcome,"sogou.") > 0 or InStr(vcome,"yahoo.") > 0 or InStr(vcome,"youdao.") > 0 or InStr(vcome,"search.live") > 0 or InStr(vcome,"search=") > 0 or InStr(vcome,"zhongsou.") > 0 then
	if InStr(vcome,"google.") > 0  or InStr(vcome,"youdao.") > 0 or InStr(vcome,"yahoo.") > 0 or InStr(vcome,"search.live") > 0 or InStr(vcome,"ie=utf-8") > 0 then
		if InStr(vcome,"google.") > 0 then
			vcomen = split(split(vcome,"q=")(1),"&")
		elseif InStr(vcome,"youdao.") > 0 and InStr(vcome,"search?q=") > 0 then
			vcomen = split(split(vcome,"?q=")(1),"&")
		elseif InStr(vcome,"youdao.") > 0 and InStr(vcome,"&q=") > 0 then
			vcomen = split(split(vcome,"&q=")(1),"&")
		elseif InStr(vcome,"yahoo.") > 0 then
			vcomen = split(split(vcome,"p=")(1),"&")
		elseif InStr(vcome,"search.live") > 0 then
			vcomen = split(split(vcome,"q=")(1),"&")
		elseif InStr(vcome,"baidu.") > 0 and InStr(vcome,"word=") > 0 then
			vcomen = split(split(vcome,"word=")(1),"&")
		elseif InStr(vcome,"baidu.") > 0 and InStr(vcome,"wd=") > 0 then
			vcomen = split(split(vcome,"wd=")(1),"&")
		else
			vcomen = split(split(vcome,"word=")(1),"&") 
		end if
		vkeyword = UTF2GB(Tj_keyword(vcomen(0)))	
	else
		if InStr(vcome,"baidu.") > 0 and InStr(vcome,"word=") > 0 then 
			vcomen = split(split(vcome,"word=")(1),"&")
		elseif InStr(vcome,"baidu.")>0 and InStr(vcome,"wd=")>0 then 
		   vcomen = split(split(vcome,"wd=")(1),"&")
		elseif InStr(vcome,"soso.")>0 then
		   vcomen = split(split(vcome,"w=")(1),"&")
		elseif InStr(vcome,"sogou.")>0 then
		   vcomen = split(split(vcome,"query=")(1),"&")
		elseif InStr(vcome,"search=")>0 then
		   vcomen = split(split(vcome,"search=")(1),"&")
		end if
		vkeyword = URLDecode(Tj_keyword(vcomen(0)))
	end if
else
	vkeyword = "未知"
end if
'屏幕分辨率获取失败时
if Cstr(screeninfo) = "[object GeneralObject]×[object GeneralObject]" then screeninfo = "获取失败"
'无来路时，判定为直接用网址打开本站
if vcome = "" then vcome = "从浏览器输入网址打开"
'记录访客及更新在线状态
sql = "select * from Tj_online where [username] = '"&Tj_user&"'"
set Tjserver = server.createobject("adodb.recordset")
Tjserver.Open sql,conn,1,3
if Tjserver.eof then
	Tjserver.addnew
	Tjserver("username") = Tj_user
	Tjserver("vcome") = vcome
	Tjserver("vcheck") = vcheck
	Tjserver("vkeyword") = vkeyword
end if
Tjserver("lastip") = ip
Tjserver("lasttime") = now()
Tjserver("vpage") = vpage
Tjserver("vwhere") = vwhere
Tjserver("vwheref") = vwheref
Tjserver("browser") = browser
Tjserver("systemer") = systemer
Tjserver("screeninfo") = screeninfo
Tjserver.update
Tjserver.close
set Tjserver = nothing
Function URLDecode(enStr)
	dim deStr
	dim c,i,v
	deStr=""
	for i = 1 to len(enStr)
		c = Mid(enStr,i,1)
		if c = "%" then
			v = eval("&h"+Mid(enStr,i+1,2))
			if v < 128 then
				deStr=deStr&chr(v)
				i = i + 2
			else
				if isvalidhex(mid(enstr,i,3)) then
					if isvalidhex(mid(enstr,i+3,3)) then
			 			v = eval("&h"+Mid(enStr,i+1,2)+Mid(enStr,i+4,2))
			 			deStr = deStr&chr(v)
			 			i = i + 5
					else
						v = eval("&h"+Mid(enStr,i+1,2)+cstr(hex(asc(Mid(enStr,i+3,1)))))
						deStr = deStr&chr(v)
						i = i + 3 
					end if 
				else 
					destr = destr&c
				end if
		  	end if
		else
			if c = "+" then
				deStr = deStr&" "
			else
				deStr = deStr&c
			end if
		end if
	next
	URLDecode=deStr
end function
function isvalidhex(str)
	dim c
	isvalidhex = true
	str = ucase(str)
	if len(str) <> 3 then isvalidhex = false:exit function
	if left(str,1) <> "%" then isvalidhex = false:exit function
	c = mid(str,2,1)
	if not (((c>="0") and (c<="9")) or ((c>="A") and (c<="Z"))) then isvalidhex = false:exit function
	c=mid(str,3,1)
	if not (((c>="0") and (c<="9")) or ((c>="A") and (c<="Z"))) then isvalidhex = false:exit function
end function
function Tj_keyword(Tj_str)
	Tj_strt = Replace(Tj_str, "?", "？")
	Tj_strt = Replace(Tj_str, "#", "＃")
	Tj_strt = Replace(Tj_str, "&", "＆")
	Tj_strt = Replace(Tj_str, "<", "&lt;")
	Tj_strt = Replace(Tj_str, ">", "&gt;")
	Tj_strt = Replace(Tj_str, Chr(13), "<br>")
	Tj_strt = Replace(Tj_str, Chr(32), "&nbsp;")
	Tj_strt = Replace(Tj_str, Chr(34), "&quot;")
	Tj_strt = Replace(Tj_str, Chr(39), "&#39")
	Tj_keyword = Tj_strt
end function
function UTF2GB(UTFStr)
	for Dig = 1 to len(UTFStr)
	if mid(UTFStr,Dig,1) = "%" then
	if len(UTFStr) >= Dig + 8 then
		GBStr = GBStr & ConvChinese(mid(UTFStr,Dig,9))
		Dig = Dig + 8
	else
		GBStr = GBStr & mid(UTFStr,Dig,1)
	end if
	else
		GBStr = GBStr & mid(UTFStr,Dig,1)
	end if
	next
		UTF2GB = GBStr
end function
function ConvChinese(x) 
	A = split(mid(x,2),"%") 
	i = 0 
	j = 0 
	for i = 0 to ubound(A) 
		A(i) = c16to2(A(i)) 
	next
	for i = 0 to ubound(A) - 1
		DigS = instr(A(i),"0")
		Unicode = ""
		for j = 1 to DigS - 1
			if j = 1 then
				A(i) = right(A(i),len(A(i)) - DigS)
				Unicode = Unicode & A(i)
			else
				i = i + 1
				A(i) = right(A(i),len(A(i)) - 2)
				Unicode = Unicode & A(i)
			end if
		next
		if len(c2to16(Unicode)) = 4 then 
			ConvChinese = ConvChinese & chrw(int("&H" & c2to16(Unicode))) 
		else 
			ConvChinese = ConvChinese & chr(int("&H" & c2to16(Unicode))) 
		end if 
	next 
end function 
function c2to16(x) 
	i = 1 
	for i = 1 to len(x) step 4 
		c2to16 = c2to16 & hex(c2to10(mid(x,i,4))) 
	next 
end function 
function c2to10(x) 
	c2to10 = 0 
	if x = "0" then exit function 
	i = 0 
	for i = 0 to len(x) - 1 
		if mid(x,len(x)-i,1) = "1" then c2to10 = c2to10 + 2^(i) 
	next 
end function 
function c16to2(x) 
	i = 0 
	for i = 1 to len(trim(x)) 
		tempstr = c10to2(cint(int("&h" & mid(x,i,1)))) 
	do while len(tempstr)<4 
		tempstr = "0" & tempstr 
	loop 
		c16to2 = c16to2 & tempstr 
	next 
end function 
function c10to2(x) 
	mysign = sgn(x) 
	x = abs(x) 
	DigS = 1 
	do
		if x<2^DigS then 
			exit do 
		else 
			DigS=DigS+1 
		end if 
	loop 
	tempnum = x 
	i = 0 
	for i = DigS to 1 step - 1 
	if tempnum >= 2^(i-1) then 
		tempnum = tempnum - 2^(i-1) 
		c10to2 = c10to2 & "1" 
	else 
		c10to2 = c10to2 & "0" 
	end if 
	next 
	if mysign = -1 then c10to2 = "-" & c10to2 
end function 
%>