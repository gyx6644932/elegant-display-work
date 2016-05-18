<!--#include file="../inc/conn.asp"-->
<%
Dim tempstr,Oper,date1,date2,ipsum,online,tdayip,ydayip,montip,yearip,allday
Oper = request("Oper")
date1 = request("date1")
date2 = request("date2")
if Oper = "del" then
	Conn.Execute("Delete From Tj_online Where lasttime between '"&date1&"' and '"&date2&"'")
	Conn.Execute("Delete From Tj_stat Where view_date between '"&date1&"' and '"&date2&"'")
	Response.Redirect("stat.asp")
	Response.End()
elseif Oper = "delall" then
	Conn.Execute("Delete From Tj_online")
	Conn.Execute("Delete From Tj_stat")
	Response.Redirect("stat.asp")
	Response.End()
end if
'总IP访问量
ipsum = conn.execute("Select sum(view_number) from Tj_stat" )(0)
if ipsum = "" or isnull(ipsum) then ipsum = 0
'总点击量
chksum = conn.execute("Select count(id) from Tj_online" )(0)
'在线人数（20分钟无动作视为离线）
online = conn.execute("select count(id) from Tj_online where [lasttime] >= DateAdd(ss,0 - 1200,Getdate())")(0)
'今日IP访问量
tdayip = conn.execute("Select sum(view_number) from Tj_stat where CONVERT(varchar(100),[view_date],112) = CONVERT(varchar(100),GetDate(),112)")(0)
if tdayip = "" or isnull(tdayip) then tdayip = 0
'昨天IP访问量
ydayip = conn.execute("Select sum(view_number) from Tj_stat where CONVERT(varchar(100),[view_date],112) = CONVERT(varchar(100),DateAdd(dd,-1,Getdate()),112)")(0)
if ydayip = "" or isnull(ydayip) then ydayip = 0
'本月IP访问量
montip = conn.execute("Select sum(view_number) from Tj_stat where view_date between DATEADD(mm,DATEDIFF(mm,0,Getdate()),0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(m,0,getdate())+1,0))")(0)
if montip = "" or isnull(montip) then montip = 0
'今年IP访问量
yearip = conn.execute("Select sum(view_number) from Tj_stat where view_date between DATEADD(yy,DATEDIFF(yy,0,Getdate()),0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate())+1,0))")(0)
if yearip = "" or isnull(yearip) then yearip = 0
'统计天数
allday = conn.execute("Select count(view_number) from Tj_stat")(0)
if allday = 0 then allday = 1
%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<!--[if !mso]>
<style>
v\:*         { behavior: url(#default#VML) }
o\:*         { behavior: url(#default#VML) }
.shape       { behavior: url(#default#VML) }
</style>
<![endif]-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/Admin.css" rel="stylesheet" type="text/css">
<script language="javascript" type="text/javascript" src="../datepicker/WdatePicker.js"></script>
<script language="javascript">function delcfm(){if(!confirm("确认要删除？")){window.event.returnValue = false;}}</script>
</head>
<body>
<table cellpadding="0" cellspacing="1" class="border">
  <tr><th colspan="10">总体数据</th></tr>
  <tr class="tdbg2">
    <td>统计持续天数</td>
    <td>昨日IP访问量</td>
    <td>今日IP访问量</td>
    <td>本月IP访问量</td>
    <td>今年IP访问量</td>
    <td>总IP访问量</td>
    <td>平均IP访问量</td>
    <td>网站总点击量</td>
    <td>平均点击量</td>
    <td>当前在线访客</td>
  </tr>
  <tr class="tdbg">
    <td class="t-center h25"><strong><%=allday%></strong> 天</td>
    <td class="t-center"><strong><%=ydayip%></strong> 个</td>
    <td class="t-center"><strong><%=tdayip%></strong> 个</td>
    <td class="t-center"><strong><%=montip%></strong> 个</td>
    <td class="t-center"><strong><%=yearip%></strong> 个</td>
    <td class="t-center"><strong><%=ipsum%></strong> 个</td>
    <td class="t-center"><strong><%=Clng(Clng(ipsum) / Clng(allday))%></strong> 个/天</td>
    <td class="t-center"><strong><%=chksum%></strong> 次</td>
    <td class="t-center"><strong><%=Clng(Clng(chksum) / Clng(allday))%></strong> 次/天</td>
    <td class="t-center"><strong><%=online%></strong> 人</td>
  </tr>
  <tr class="tdbg"><td colspan="10" class="t-left h25"><span class="red">（注）</span><strong>IP访问量</strong>：同一个访客（同一台计算机、同一个公司单位内所有计算机、同一个网吧内所有计算机只算1）在当天不限次数访问网站只计算为1。 </td></tr>
  <tr class="tdbg"><td colspan="10" class="t-left h25"><span class="red">（注）</span><strong>点击量</strong>：打开网站后直到关闭所有页面的过程计算为1（浏览多页面不增加），重新打开网站增加1。</td></tr>
  <tr class="tdbg"><td colspan="10" class="t-left h25"><span class="red">（注）</span>后台页面不计算在内，不做统计。</td></tr>
  <form method="post">
  <input type="hidden" name="Oper" value="del">
  <tr class="tdbg2"><td colspan="10">从 <input name="date1" id="date1" value="<%=date()-395%>" type="text" onFocus="WdatePicker({dateFmt:'yyyy-M-d',maxDate:'#F{$dp.$D(\'date2\',{d:-1})}',isShowClear:false,readOnly:true})" style="width:85px;" readonly> 到 <input name="date2" id="date2" value="<%=date()-30%>" type="text" onFocus="WdatePicker({dateFmt:'yyyy-M-d',minDate:'#F{$dp.$D(\'date1\',{d:1})}',maxDate:'%y-%M-%d',isShowClear:false,readOnly:true})" style="width:85px;" readonly> <input type="submit" class="bt" value="删除" onClick="delcfm()"> <a href="stat.asp?Oper=delall" onClick="delcfm()">全部清空</a></td></tr>
  </form>
</table>
<br>
<%
ipdate1 = Request("ipdate1")
ipdate2 = Request("ipdate2")
%>
<table cellpadding="0" cellspacing="1" class="border">
  <tr><th colspan="2"><a name="ip"></a>IP 统 计</th></tr>
  <form method="get" action="stat.asp#ip">
  <tr class="tdbg">
    <td colspan="2" class="t-center"><a href="?action=ipall#ip">全部</a>　<a href="?action=iptoday#ip">今天</a>　<a href="?action=ipyesterday#ip">昨天</a>　<a href="?action=ipthismonth#ip">本月</a>　<a href="?action=iplastmonth#ip">上月</a>　<a href="?action=ipthisyear#ip">今年</a>　<a href="?action=iplastyear#ip">去年</a>　　选择范围：从 <input type="text" name="ipdate1" id="ipdate1" value="<%=ipdate1%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',maxDate:'#F{$dp.$D(\'ipdate2\',{d:-1})}',readOnly:true})" style="width:85px;" readonly> 到 <input type="text" name="ipdate2" id="ipdate2" value="<%=ipdate2%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',minDate:'#F{$dp.$D(\'ipdate1\',{d:1})}',maxDate:'%y-%M-%d',readOnly:true})" style="width:85px;" readonly> <input type="submit" class="bt" value="筛选"></td>
  </tr>
  </form>
  <tr class="tdbg">
    <td valign="top">
      <table cellpadding="0" cellspacing="1" style="width:100%; background-color:#ABDEEF;">
        <tr class="tdbg2">
          <td>日期</td>
          <td>IP访问量</td>
        </tr>
		<%
		action = request("action")
		If ipdate1 <> "" Or ipdate2 <> "" Then action = ""
		sql = "select * from Tj_stat where view_date is not null"
		select case action
			case "ipall"
			 	sql = sql
			case "iptoday"
				sql = sql&" and CONVERT(varchar(100),[view_date],112) = CONVERT(varchar(100),GetDate(),112)"
			case "ipyesterday"
				sql = sql&" and CONVERT(varchar(100),[view_date],112) = CONVERT(varchar(100),DateAdd(dd,-1,Getdate()),112)"
			case "ipthismonth"
				sql = sql&" and (view_date between DATEADD(mm,DATEDIFF(mm,0,Getdate()),0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate())+1,0)))"
			case "iplastmonth"
				sql = sql&" and (view_date between DATEADD(mm,DATEDIFF(mm,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate()),0)))"
			case "ipthisyear"
				sql = sql&" and (view_date between DATEADD(yy,DATEDIFF(yy,0,Getdate()),0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate())+1,0)))"
			case "iplastyear"
				sql = sql&" and (view_date between DATEADD(yy,DATEDIFF(yy,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate()),0)))"
	    end select	
		if ipdate1 <> "" then sql = sql&" and view_date >= '"&ipdate1&"'"
		if ipdate2 <> "" then sql = sql&" and view_date <= '"&ipdate2&"'"	  
		sql = sql&" order by view_date desc"
        set Trs = server.createobject("adodb.recordset")
        Trs.open sql,conn,1,1
		dim maxperpage,currentpage,allpage,totalput,i,heji
		maxperpage = 13
		dim total(13,1)
		total(0,1) = "#FF0000,1.5,1,2,IP统计"
		if Trs.eof then
		currentpage = 1
		allpage = 1
		totalput = 0
		for ii = 1 to maxperpage
			total(ii,0) = ""
			total(ii,1) = 0
		next
		%>
        <tr class="tdbg">
          <td colspan="2" class="t-center h25">暂无数据</td>
        </tr>
        <%
		else
		Trs.pagesize = maxperpage
		currentpage = trim(request("pageid"))
		if currentpage = "" then currentpage = 1
		allpage = Trs.pagecount
		totalput = Trs.recordcount
		i=0
		heji = 0
		call pagelist1(currentpage,maxperpage,allpage)
        do while i < maxperpage and not Trs.eof%>
        <tr class="tdbg">
          <td class="t-center h25"><%=Trs("view_date")%></td>
          <td class="t-center"><%=Trs("view_number")%></td>
        </tr>
        <%
		heji = heji + Trs("view_number")
		total(i + 1,0) = month(Trs("view_date"))&"/"&day(Trs("view_date"))
		total(i + 1,1) = Trs("view_number")
		Trs.movenext
		i = i + 1
		loop
		%>
        <tr align="center" class="tdbg2">
          <td><font color=red>当前合计</font></td>
          <td><font color=red><%=heji%></font></td>
        </tr>
        <%
		end if
		Trs.close
		set Trs = nothing
		%>
      </table>
    </td>
    <td style="width:80%;" align="center"><div style="position:relative;width:600px;height:250px;text-align:center"><%call table2(total,0,20,514,200,1)%></div></td>
  </tr>
  <tr class="tdbg2"><td colspan="2"><%call pagelist2(currentpage,allpage,totalput,"action="&action&"&ipdate1="&ipdate1&"&ipdate2="&ipdate2,"ip")%></td></tr>
</table>
<br>
<%
redate1 = Request("redate1")
redate2 = Request("redate2")
%>
<table cellpadding="0" cellspacing="1" class="border">
  <tr><th colspan="10"><a name="re"></a>详细记录</th></tr>
  <form method="get" action="stat.asp#re">
  <tr class="tdbg">
    <td colspan="10" class="t-center"><a href="?action=reall#re">全部</a>　<a href="?action=retoday#re">今天</a>　<a href="?action=reyesterday#re">昨天</a>　<a href="?action=rethismonth#re">本月</a>　<a href="?action=relastmonth#re">上月</a>　<a href="?action=rethisyear#re">今年</a>　<a href="?action=relastyear#re">去年</a>　　选择范围：从 <input type="text" name="redate1" id="redate1" value="<%=redate1%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',maxDate:'#F{$dp.$D(\'redate2\',{d:-1})}',readOnly:true})" style="width:85px;" readonly> 到 <input type="text" name="redate2" id="redate2" value="<%=redate2%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',minDate:'#F{$dp.$D(\'redate1\',{d:1})}',maxDate:'%y-%M-%d',readOnly:true})" style="width:85px;" readonly> <input type="submit" class="bt" value="筛选"></td>
  </tr>
  </form>
  <tr class="tdbg2">
    <td>访问时间</td>
    <td>所在地区</td>
    <td>使用网络</td>
    <td>搜索引擎</td>
    <td>用此关键字搜索</td>
    <td>从此页面找到本站<font color="red">(点击查看)</font></td>
    <td>本站被访问页<font color="red">(点击查看)</font></td>
    <td>使用的系统</td>
    <td>使用的浏览器</td>
    <td>屏幕分辨率</td>
  </tr>
  <%
  action = request("action")
  If redate1 <> "" Or redate2 <> "" Then action = ""
  sql="select * from Tj_online where id is not null"
  select case action
	  case "reall"
		  sql = sql
	  case "retoday"
		  sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),GetDate(),112)"
	  case "reyesterday"
		  sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),DateAdd(dd,-1,Getdate()),112)"
	  case "rethismonth"
	      sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate()),0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate())+1,0)))"
	  case "relastmonth"
	      sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate()),0)))"
	  case "rethisyear"
	      sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate()),0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate())+1,0)))"
	  case "relastyear"
		  sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate()),0)))"
  end select	
  if redate1 <> "" then sql = sql&" and lasttime >= '"&redate1&"'"
  if redate2 <> "" then sql = sql&" and lasttime <= '"&todate2&"'"
  sql=sql&" order by id desc"
  set Trs = server.createobject("adodb.recordset")
  Trs.open sql,conn,1,1
  maxperpage = 20 
  Trs.pagesize = maxperpage
  currentpage = trim(request("pageid"))	
  if currentpage = "" then currentpage = 1
  allpage = Trs.pagecount 
  totalput = Trs.recordcount 
  i=0
  call pagelist1(currentpage,maxperpage,allpage)	   	  
  do while i < maxperpage and not Trs.eof
  %>
  <tr class="tdbg">
    <td class="t-center h25"><%=Trs("lasttime")%></td>
    <td class="t-center"><%=Trs("vwhere")%></td>
    <td class="t-center"><%=Trs("vwheref")%></td>
    <td class="t-center"><%=Trs("vcheck")%></td>
    <td class="t-center"><%=Trs("vkeyword")%></td>
    <td style="word-wrap:break-word;word-break:break-all;text-align:left;"><a href="<%if Cstr(Trs("vcome")) = "从浏览器输入网址打开" then%>about:blank<%else%><%=Trs("vcome")%><%end if%>" target="_blank" ><%=Left(Trs("vcome"),38)&"..."%></a></td>
    <td style="word-wrap:break-word;word-break:break-all;text-align:left;"><a href="<%=Trs("vpage")%>" target="_blank" ><%=Left(Trs("vpage"),38)&"..."%></a></td>
    <td class="t-center"><%=Trs("systemer")%></td>
    <td class="t-center"><%=Trs("browser")%></td>
    <td class="t-center"><%=Trs("screeninfo")%></td>
  </tr>
  <%
  Trs.movenext
  i = i + 1
  loop
  Trs.close
  set Trs = nothing
  %>
  <tr class="tdbg2"><td colspan="10"><%call pagelist2(currentpage,allpage,totalput,"action="&action&"&redate1="&redate1&"&redate2="&redate2,"re")%></td></tr>
</table>
<br>
<%
checkerdate1 = Request.Form("checkerdate1")
checkerdate2 = Request.Form("checkerdate2")
%>
<table cellpadding="0" cellspacing="1" class="border">
  <tr><th><a name="checker"></a>用户使用搜索引擎统计</th></tr>
  <form method="get" action="stat.asp#checker">
  <tr class="tdbg">
    <td colspan="10" class="t-center h25"><a href="?action=soall#checker">全部</a>　<a href="?action=sotoday#checker">今天</a>　<a href="?action=soyesterday#checker">昨天</a>　<a href="?action=sothismonth#checker">本月</a>　<a href="?action=solastmonth#checker">上月</a>　<a href="?action=sothisyear#checker">今年</a>　<a href="?action=solastyear#checker">去年</a>　　选择范围：从 <input type="text" name="checkerdate1" id="checkerdate1" value="<%=checkerdate1%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',maxDate:'#F{$dp.$D(\'checkerdate2\',{d:-1})}',readOnly:true})" style="width:85px;" readonly> 到 <input type="text" name="checkerdate2" id="checkerdate2" value="<%=checkerdate2%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',minDate:'#F{$dp.$D(\'checkerdate1\',{d:1})}',maxDate:'%y-%M-%d',readOnly:true})" style="width:85px;" readonly> <input type="submit" class="bt" value="筛选"></td>
  </tr>
  </form>
  <tr class="tdbg">
    <td align="center" style="padding-top:20px; padding-right:20px; padding-bottom:20px;">
      <table width="100%"  border="0" cellpadding="0" cellspacing="0">
        <%
		action = request("action")
		If checkerdate1 <> "" Or checkerdate2 <> "" Then action = ""
		sql = "select vcheck,count(id) as allcheck from Tj_online where id is not null"
		select case action
			case "soall"
				sql = sql
			case "sotoday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),GetDate(),112)"
			case "soyesterday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),DateAdd(dd,-1,Getdate()),112)"
			case "sothismonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate()),0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate())+1,0)))"
			case "solastmonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate()),0)))"
			case "sothisyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate()),0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate())+1,0)))"
			case "solastyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate()),0)))"
	    end select
		if checkerdate1 <> "" then sql = sql&" and lasttime >= '"&checkerdate1&"'"
		if checkerdate2 <> "" then sql = sql&" and lasttime <= '"&checkerdate2&"'"  
		sql = sql&" group by vcheck order by count(id) DESC"
		set Trs = server.createobject("adodb.recordset")
        Trs.Open sql,conn,1,1 
		if Trs.eof then msg = "<font color=red>没数据</font>"
		i = 1
        Do While Not Trs.eof and i < 21
			picwidth = Trs("allcheck")
			if i = 1 then maxpicwidth = picwidth 
			if maxpicwidth < 600 then
				picwidth = Trs("allcheck")
			elseif maxpicwidth >= 600 and maxpicwidth <= 6000 then
				picwidth = int(picwidth/10)
			elseif maxpicwidth >= 6000 and maxpicwidth <= 60000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 60000 and maxpicwidth <= 600000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 600000 and maxpicwidth <= 6000000 then
				picwidth = int(picwidth/1000)
			else
				picwidth = int(picwidth/10000)
			end if      
			%>
        <tr>
          <td style="width:18%; height:25px;text-align:right;"><%=Trs("vcheck")%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;text-align:left;padding-left:0px;"><img src="images/state.jpg" width="<%=picwidth%>" height="17" align="absmiddle">&nbsp;<%=Trs("allcheck")%></td>
        </tr>
			<% 
			Trs.movenext
			i = i + 1
        Loop
        Trs.close
		set Trs = nothing
		%>
        <tr>
          <td align="right"><%=msg%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;BORDER-bottom: #D3D3D3 1px solid;"><img src="images/tu.gif" width="1" height="1"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<%
keydate1 = Request.Form("keydate1")
keydate2 = Request.Form("keydate2")
%>
<table cellpadding="2" cellspacing="1"class="border">
  <tr><th><a name="key"></a>用户在搜索引擎中找到本站用的关键词统计</th></tr>
  <form method="get" action="stat.asp#key">
  <tr class="tdbg">
    <td colspan="10" class="t-center h25"><a href="?action=keyall#key">全部</a>　<a href="?action=keytoday#key">今天</a>　<a href="?action=keyyesterday#key">昨天</a>　<a href="?action=keythismonth#key">本月</a>　<a href="?action=keylastmonth#key">上月</a>　<a href="?action=keythisyear#key">今年</a>　<a href="?action=keylastyear#key">去年</a>　　选择范围：从 <input type="text" name="keydate1" id="keydate1" value="<%=keydate1%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',maxDate:'#F{$dp.$D(\'keydate2\',{d:-1})}',readOnly:true})" style="width:85px;" readonly> 到 <input type="text" name="keydate2" id="keydate2" value="<%=keydate2%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',minDate:'#F{$dp.$D(\'keydate1\',{d:1})}',maxDate:'%y-%M-%d',readOnly:true})" style="width:85px;" readonly> <input type="submit" class="bt" value="筛选"></td>
  </tr>
  </form>
  <tr class="tdbg">
    <td align="center" style="padding-top:20px; padding-right:20px; padding-bottom:20px;">
      <table width="100%"  border="0" cellpadding="0" cellspacing="0">
        <%
		action = request("action")
		If keydate1 <> "" Or keydate2 <> "" Then action = ""
		sql = "select vkeyword,count(id) as allkey from Tj_online where id is not null"
		select case action
			case "keyall"
				sql = sql
			case "keytoday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),GetDate(),112)"
			case "keyyesterday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),DateAdd(dd,-1,Getdate()),112)"
			case "keythismonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate()),0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate())+1,0)))"
			case "keylastmonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate()),0)))"
			case "keythisyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate()),0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate())+1,0)))"
			case "keylastyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate()),0)))"
	    end select
		if keydate1 <> "" then sql = sql&" and lasttime >= '"&keydate1&"'"
		if keydate2 <> "" then sql = sql&" and lasttime <= '"&keydate2&"'"
		sql = sql&" group by vkeyword order by count(id) DESC"
		set Trs = server.createobject("adodb.recordset")
        Trs.Open sql,conn,1,1
		if Trs.eof then msg = "<font color=red>没数据</font>"
		i = 1
        Do While Not Trs.eof and i < 51
			picwidth = Trs("allkey")
			if i = 1 then maxpicwidth = picwidth 
			if maxpicwidth < 600 then
				picwidth = Trs("allkey")
			elseif maxpicwidth >= 600 and maxpicwidth <= 6000 then
				picwidth = int(picwidth/10)
			elseif maxpicwidth >= 6000 and maxpicwidth <= 60000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 60000 and maxpicwidth <= 600000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 600000 and maxpicwidth <= 6000000 then
				picwidth = int(picwidth/1000)
			else
				picwidth = int(picwidth/10000)
			end if      
		%>
        <tr>
          <td style="width:18%;height:25px;text-align:right;"><%=Trs("vkeyword")%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;text-align:left;padding-left:0px;"><img src="images/state.jpg" width="<%=picwidth%>" height="17" align="absmiddle">&nbsp;<%=Trs("allkey")%></td>
        </tr>
		<% 
			Trs.movenext
			i = i + 1
        Loop
        Trs.close
		set Trs = nothing
		%>
        <tr>
          <td align="right"><%=msg%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;BORDER-bottom: #D3D3D3 1px solid;"><img src="images/tu.gif" width="1" height="1"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<%
sysdate1 = Request.Form("sysdate1")
sysdate2 = Request.Form("sysdate2")
%>
<table cellpadding="2" cellspacing="1"class="border">
  <tr><th><a name="sys"></a>用户计算机操作系统统计</th></tr>
  <form method="get" action="stat.asp#sys">
  <tr class="tdbg">
    <td colspan="10" class="t-center h25"><a href="?action=sysall#sys">全部</a>　<a href="?action=systoday#sys">今天</a>　<a href="?action=sysyesterday#sys">昨天</a>　<a href="?action=systhismonth#sys">本月</a>　<a href="?action=syslastmonth#sys">上月</a>　<a href="?action=systhisyear#sys">今年</a>　<a href="?action=syslastyear#sys">去年</a>　　选择范围：从 <input type="text" name="sysdate1" id="sysdate1" value="<%=sysdate1%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',maxDate:'#F{$dp.$D(\'sysdate2\',{d:-1})}',readOnly:true})" style="width:85px;" readonly> 到 <input type="text" name="sysdate2" id="sysdate2" value="<%=sysdate2%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',minDate:'#F{$dp.$D(\'sysdate1\',{d:1})}',maxDate:'%y-%M-%d',readOnly:true})" style="width:85px;" readonly> <input type="submit" class="bt" value="筛选"></td>
  </tr>
  </form>
  <tr class="tdbg">
    <td align="center" style="padding-top:20px; padding-right:20px; padding-bottom:20px;">
      <table width="100%"  border="0" cellpadding="0" cellspacing="0">
        <%
		action = request("action")
		If sysdate1 <> "" Or sysdate2 <> "" Then action = ""
		sql = "select systemer,count(id) as allsys from Tj_online where id is not null"
		select case action
			case "sysall"
				sql = sql
			case "systoday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),GetDate(),112)"
			case "sysyesterday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),DateAdd(dd,-1,Getdate()),112)"
			case "systhismonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate()),0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate())+1,0)))"
			case "syslastmonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate()),0)))"
			case "systhisyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate()),0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate())+1,0)))"
			case "syslastyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate()),0)))"
	    end select
		if sysdate1 <> "" then sql = sql&" and lasttime >= '"&sysdate1&"'"
		if sysdate2 <> "" then sql = sql&" and lasttime <= '"&sysdate2&"'"
		sql = sql&" group by systemer order by count(id) DESC"
		set Trs = server.createobject("adodb.recordset")
        Trs.Open sql,conn,1,1
		if Trs.eof then msg = "<font color=red>没数据</font>"
		i = 1
        Do While Not Trs.eof and i < 51
			picwidth = Trs("allsys")
			if i = 1 then maxpicwidth = picwidth 
			if maxpicwidth < 600 then
				picwidth = Trs("allsys")
			elseif maxpicwidth >= 600 and maxpicwidth <= 6000 then
				picwidth = int(picwidth/10)
			elseif maxpicwidth >= 6000 and maxpicwidth <= 60000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 60000 and maxpicwidth <= 600000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 600000 and maxpicwidth <= 6000000 then
				picwidth = int(picwidth/1000)
			else
				picwidth = int(picwidth/10000)
			end if      
		%>
        <tr>
          <td style="width:18%;height:25px;text-align:right;"><%=Trs("systemer")%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;text-align:left;padding-left:0px;"><img src="images/state.jpg" width="<%=picwidth%>" height="17" align="absmiddle">&nbsp;<%=Trs("allsys")%></td>
        </tr>
		<% 
			Trs.movenext
			i = i + 1
        Loop
        Trs.close
		set Trs = nothing
		%>
        <tr>
          <td align="right"><%=msg%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;BORDER-bottom: #D3D3D3 1px solid;"><img src="images/tu.gif" width="1" height="1"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<%
adsl1 = Request.Form("adsl1")
adsl2 = Request.Form("adsl2")
%>
<table cellpadding="2" cellspacing="1"class="border">
  <tr><th><a name="sys"></a>用户使用的网络情况统计</th></tr>
  <form method="get" action="stat.asp#adsl">
  <tr class="tdbg">
    <td colspan="10" class="t-center h25"><a href="?action=adslall#adsl">全部</a>　<a href="?action=adsltoday#adsl">今天</a>　<a href="?action=adslyesterday#adsl">昨天</a>　<a href="?action=adslthismonth#adsl">本月</a>　<a href="?action=adsllastmonth#adsl">上月</a>　<a href="?action=adslthisyear#adsl">今年</a>　<a href="?action=adsllastyear#adsl">去年</a>　　选择范围：从 <input type="text" name="adsl1" id="adsl1" value="<%=adsl1%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',maxDate:'#F{$dp.$D(\'adsl2\',{d:-1})}',readOnly:true})" style="width:85px;" readonly> 到 <input type="text" name="adsl2" id="adsl2" value="<%=adsl2%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',minDate:'#F{$dp.$D(\'adsl1\',{d:1})}',maxDate:'%y-%M-%d',readOnly:true})" style="width:85px;" readonly> <input type="submit" class="bt" value="筛选"></td>
  </tr>
  </form>
  <tr class="tdbg">
    <td align="center" style="padding-top:20px; padding-right:20px; padding-bottom:20px;">
      <table width="100%"  border="0" cellpadding="0" cellspacing="0">
        <%
		action = request("action")
		If adsl1 <> "" Or adsl2 <> "" Then action = ""
		sql = "select vwheref,count(id) as alladsl from Tj_online where id is not null"
		select case action
			case "adslall"
				sql = sql
			case "adsltoday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),GetDate(),112)"
			case "adslyesterday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),DateAdd(dd,-1,Getdate()),112)"
			case "adslthismonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate()),0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate())+1,0)))"
			case "adsllastmonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate()),0)))"
			case "adslthisyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate()),0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate())+1,0)))"
			case "adsllastyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate()),0)))"
	    end select
		if adsl1 <> "" then sql = sql&" and lasttime >= '"&adsl1&"'"
		if adsl2 <> "" then sql = sql&" and lasttime <= '"&adsl2&"'"
		sql = sql&" group by vwheref order by count(id) DESC"
		set Trs = server.createobject("adodb.recordset")
        Trs.Open sql,conn,1,1
		if Trs.eof then msg = "<font color=red>没数据</font>"
		i = 1
        Do While Not Trs.eof and i < 51
			picwidth = Trs("alladsl")
			if i = 1 then maxpicwidth = picwidth 
			if maxpicwidth < 600 then
				picwidth = Trs("alladsl")
			elseif maxpicwidth >= 600 and maxpicwidth <= 6000 then
				picwidth = int(picwidth/10)
			elseif maxpicwidth >= 6000 and maxpicwidth <= 60000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 60000 and maxpicwidth <= 600000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 600000 and maxpicwidth <= 6000000 then
				picwidth = int(picwidth/1000)
			else
				picwidth = int(picwidth/10000)
			end if      
		%>
        <tr>
          <td style="width:18%;height:25px;text-align:right;"><%=Trs("vwheref")%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;text-align:left;padding-left:0px;"><img src="images/state.jpg" width="<%=picwidth%>" height="17" align="absmiddle">&nbsp;<%=Trs("alladsl")%></td>
        </tr>
		<% 
			Trs.movenext
			i = i + 1
        Loop
        Trs.close
		set Trs = nothing
		%>
        <tr>
          <td align="right"><%=msg%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;BORDER-bottom: #D3D3D3 1px solid;"><img src="images/tu.gif" width="1" height="1"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<%
browdate1 = Request.Form("browdate1")
browdate2 = Request.Form("browdate2")
%>
<table cellpadding="2" cellspacing="1"class="border">
  <tr><th><a name="brow"></a>用户使用浏览器统计</th></tr>
  <form method="get" action="stat.asp#brow">
  <tr class="tdbg">
    <td colspan="10" class="t-center h25"><a href="?action=browall#brow">全部</a>　<a href="?action=browtoday#brow">今天</a>　<a href="?action=browyesterday#brow">昨天</a>　<a href="?action=browthismonth#brow">本月</a>　<a href="?action=browlastmonth#brow">上月</a>　<a href="?action=browthisyear#brow">今年</a>　<a href="?action=browlastyear#brow">去年</a>　　选择范围：从 <input type="text" name="browdate1" id="browdate1" value="<%=browdate1%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',maxDate:'#F{$dp.$D(\'browdate2\',{d:-1})}',readOnly:true})" style="width:85px;" readonly> 到 <input type="text" name="browdate2" id="browdate2" value="<%=browdate2%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',minDate:'#F{$dp.$D(\'browdate1\',{d:1})}',maxDate:'%y-%M-%d',readOnly:true})" style="width:85px;" readonly> <input type="submit" class="bt" value="筛选"></td>
  </tr>
  </form>
  <tr class="tdbg">
    <td align="center" style="padding-top:20px; padding-right:20px; padding-bottom:20px;">
      <table width="100%"  border="0" cellpadding="0" cellspacing="0">
        <%
		action = request("action")
		If browdate1 <> "" Or browdate2 <> "" Then action = ""
		sql = "select browser,count(id) as allbrow from Tj_online where id is not null"
		select case action
			case "browall"
				sql = sql
			case "browtoday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),GetDate(),112)"
			case "browyesterday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),DateAdd(dd,-1,Getdate()),112)"
			case "browthismonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate()),0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate())+1,0)))"
			case "browlastmonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate()),0)))"
			case "browthisyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate()),0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate())+1,0)))"
			case "browlastyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate()),0)))"
	    end select
		if browdate1 <> "" then sql = sql&" and lasttime >= '"&browdate1&"'"
		if browdate2 <> "" then sql = sql&" and lasttime <= '"&browdate2&"'"
		sql = sql&" group by browser order by count(id) DESC"
		set Trs = server.createobject("adodb.recordset")
        Trs.Open sql,conn,1,1
		if Trs.eof then msg = "<font color=red>没数据</font>"
		i = 1
        Do While Not Trs.eof and i < 51
			picwidth = Trs("allbrow")
			if i = 1 then maxpicwidth = picwidth 
			if maxpicwidth < 600 then
				picwidth = Trs("allbrow")
			elseif maxpicwidth >= 600 and maxpicwidth <= 6000 then
				picwidth = int(picwidth/10)
			elseif maxpicwidth >= 6000 and maxpicwidth <= 60000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 60000 and maxpicwidth <= 600000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 600000 and maxpicwidth <= 6000000 then
				picwidth = int(picwidth/1000)
			else
				picwidth = int(picwidth/10000)
			end if      
		%>
        <tr>
          <td style="width:18%;height:25px;text-align:right;"><%=Trs("browser")%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;text-align:left;padding-left:0px;"><img src="images/state.jpg" width="<%=picwidth%>" height="17" align="absmiddle">&nbsp;<%=Trs("allbrow")%></td>
        </tr>
		<% 
			Trs.movenext
			i = i + 1
        Loop
        Trs.close
		set Trs = nothing
		%>
        <tr>
          <td align="right"><%=msg%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;BORDER-bottom: #D3D3D3 1px solid;"><img src="images/tu.gif" width="1" height="1"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<%
screendate1 = Request.Form("screendate1")
screendate2 = Request.Form("screendate2")
%>
<table cellpadding="2" cellspacing="1"class="border">
  <tr><th><a name="screen"></a>用户显示器屏幕分辨率统计</th></tr>
  <form method="get" action="stat.asp#screen">
  <tr class="tdbg">
    <td colspan="10" class="t-center h25"><a href="?action=screenall#screen">全部</a>　<a href="?action=screentoday#screen">今天</a>　<a href="?action=screenyesterday#screen">昨天</a>　<a href="?action=screenthismonth#screen">本月</a>　<a href="?action=screenlastmonth#screen">上月</a>　<a href="?action=screenthisyear#screen">今年</a>　<a href="?action=screenlastyear#screen">去年</a>　　选择范围：从 <input type="text" name="screendate1" id="screendate1" value="<%=screendate1%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',maxDate:'#F{$dp.$D(\'screendate2\',{d:-1})}',readOnly:true})" style="width:85px;" readonly> 到 <input type="text" name="screendate2" id="screendate2" value="<%=screendate2%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',minDate:'#F{$dp.$D(\'screendate1\',{d:1})}',maxDate:'%y-%M-%d',readOnly:true})" style="width:85px;" readonly> <input type="submit" class="bt" value="筛选"></td>
  </tr>
  </form>
  <tr class="tdbg">
    <td align="center" style="padding-top:20px; padding-right:20px; padding-bottom:20px;">
      <table width="100%"  border="0" cellpadding="0" cellspacing="0">
        <%
		action = request("action")
		If screendate1 <> "" Or screendate2 <> "" Then action = ""
		sql = "select screeninfo,count(id) as allscreen from Tj_online where id is not null"
		select case action
			case "screenall"
				sql = sql
			case "screentoday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),GetDate(),112)"
			case "screenyesterday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),DateAdd(dd,-1,Getdate()),112)"
			case "screenthismonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate()),0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate())+1,0)))"
			case "screenlastmonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate()),0)))"
			case "screenthisyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate()),0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate())+1,0)))"
			case "screenlastyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate()),0)))"
	    end select
		if screendate1 <> "" then sql = sql&" and lasttime >= '"&screendate1&"'"
		if screendate2 <> "" then sql = sql&" and lasttime <= '"&screendate2&"'"
		sql = sql&" group by screeninfo order by count(id) DESC"
		set Trs = server.createobject("adodb.recordset")
        Trs.Open sql,conn,1,1
		if Trs.eof then msg = "<font color=red>没数据</font>"
		i = 1
        Do While Not Trs.eof and i < 51
			picwidth = Trs("allscreen")
			if i = 1 then maxpicwidth = picwidth 
			if maxpicwidth < 600 then
				picwidth = Trs("allscreen")
			elseif maxpicwidth >= 600 and maxpicwidth <= 6000 then
				picwidth = int(picwidth/10)
			elseif maxpicwidth >= 6000 and maxpicwidth <= 60000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 60000 and maxpicwidth <= 600000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 600000 and maxpicwidth <= 6000000 then
				picwidth = int(picwidth/1000)
			else
				picwidth = int(picwidth/10000)
			end if      
		%>
        <tr>
          <td style="width:18%;height:25px;text-align:right;"><%=Trs("screeninfo")%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;text-align:left;padding-left:0px;"><img src="images/state.jpg" width="<%=picwidth%>" height="17" align="absmiddle">&nbsp;<%=Trs("allscreen")%></td>
        </tr>
		<% 
			Trs.movenext
			i = i + 1
        Loop
        Trs.close
		set Trs = nothing
		%>
        <tr>
          <td align="right"><%=msg%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;BORDER-bottom: #D3D3D3 1px solid;"><img src="images/tu.gif" width="1" height="1"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<%
areadate1 = Request.Form("areadate1")
areadate2 = Request.Form("areadate2")
%>
<table cellpadding="2" cellspacing="1"class="border">
  <tr><th><a name="area"></a>用户所在地区统计</th></tr>
  <form method="get" action="stat.asp#area">
  <tr class="tdbg">
    <td colspan="10" class="t-center h25"><a href="?action=areaall#area">全部</a>　<a href="?action=areatoday#area">今天</a>　<a href="?action=areayesterday#area">昨天</a>　<a href="?action=areathismonth#area">本月</a>　<a href="?action=arealastmonth#area">上月</a>　<a href="?action=areathisyear#area">今年</a>　<a href="?action=arealastyear#area">去年</a>　　选择范围：从 <input type="text" name="areadate1" id="areadate1" value="<%=areadate1%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',maxDate:'#F{$dp.$D(\'areadate2\',{d:-1})}',readOnly:true})" style="width:85px;" readonly> 到 <input type="text" name="areadate2" id="areadate2" value="<%=areadate2%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',minDate:'#F{$dp.$D(\'areadate1\',{d:1})}',maxDate:'%y-%M-%d',readOnly:true})" style="width:85px;" readonly> <input type="submit" class="bt" value="筛选"></td>
  </tr>
  </form>
  <tr class="tdbg">
    <td align="center" style="padding-top:20px; padding-right:20px; padding-bottom:20px;">
      <table width="100%"  border="0" cellpadding="0" cellspacing="0">
        <%
		action = request("action")
		If areadate1 <> "" Or areadate2 <> "" Then action = ""
		sql = "select vwhere,count(id) as allwhere from Tj_online where id is not null"
		select case action
			case "areaall"
				sql = sql
			case "areatoday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),GetDate(),112)"
			case "areayesterday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),DateAdd(dd,-1,Getdate()),112)"
			case "areathismonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate()),0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate())+1,0)))"
			case "arealastmonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate()),0)))"
			case "areathisyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate()),0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate())+1,0)))"
			case "arealastyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate()),0)))"
	    end select
		if areadate1 <> "" then sql = sql&" and lasttime >= '"&areadate1&"'"
		if areadate2 <> "" then sql = sql&" and lasttime <= '"&areadate2&"'"
		sql = sql&" group by vwhere order by count(id) DESC"
		set Trs = server.createobject("adodb.recordset")
        Trs.Open sql,conn,1,1
		if Trs.eof then msg = "<font color=red>没数据</font>"
		i = 1
        Do While Not Trs.eof and i < 51
			picwidth = Trs("allwhere")
			if i = 1 then maxpicwidth = picwidth 
			if maxpicwidth < 600 then
				picwidth = Trs("allwhere")
			elseif maxpicwidth >= 600 and maxpicwidth <= 6000 then
				picwidth = int(picwidth/10)
			elseif maxpicwidth >= 6000 and maxpicwidth <= 60000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 60000 and maxpicwidth <= 600000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 600000 and maxpicwidth <= 6000000 then
				picwidth = int(picwidth/1000)
			else
				picwidth = int(picwidth/10000)
			end if      
		%>
        <tr>
          <td style="width:18%;height:25px;text-align:right;"><%=Trs("vwhere")%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;text-align:left;padding-left:0px;"><img src="images/state.jpg" width="<%=picwidth%>" height="17" align="absmiddle">&nbsp;<%=Trs("allwhere")%></td>
        </tr>
		<% 
			Trs.movenext
			i = i + 1
        Loop
        Trs.close
		set Trs = nothing
		%>
        <tr>
          <td align="right"><%=msg%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;BORDER-bottom: #D3D3D3 1px solid;"><img src="images/tu.gif" width="1" height="1"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<%
pagendate1 = Request.Form("pagendate1")
pagendate2 = Request.Form("pagendate2")
%>
<table cellpadding="2" cellspacing="1"class="border">
  <tr><th><a name="pagen"></a>被问页面统计</th></tr>
  <form method="get" action="stat.asp#pagen">
  <tr class="tdbg">
    <td colspan="10" class="t-center h25"><a href="?action=pagenall#pagen">全部</a>　<a href="?action=pagentoday#pagen">今天</a>　<a href="?action=pagenyesterday#pagen">昨天</a>　<a href="?action=pagenthismonth#pagen">本月</a>　<a href="?action=pagenlastmonth#pagen">上月</a>　<a href="?action=pagenthisyear#pagen">今年</a>　<a href="?action=pagenlastyear#pagen">去年</a>　　选择范围：从 <input type="text" name="pagendate1" id="pagendate1" value="<%=pagendate1%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',maxDate:'#F{$dp.$D(\'pagendate2\',{d:-1})}',readOnly:true})" style="width:85px;" readonly> 到 <input type="text" name="pagendate2" id="pagendate2" value="<%=pagendate2%>" onFocus="WdatePicker({dateFmt:'yyyy-M-d',minDate:'#F{$dp.$D(\'pagendate1\',{d:1})}',maxDate:'%y-%M-%d',readOnly:true})" style="width:85px;" readonly> <input type="submit" class="bt" value="筛选"></td>
  </tr>
  </form>
  <tr class="tdbg">
    <td align="center" style="padding-top:20px; padding-right:20px; padding-bottom:20px;">
      <table width="100%"  border="0" cellpadding="0" cellspacing="0">
        <%
		action = request("action")
		If pagendate1 <> "" Or pagendate2 <> "" Then action = ""
		sql = "select vpage,count(id) as allpage from Tj_online where id is not null"
		select case action
			case "pagenall"
				sql = sql
			case "pagentoday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),GetDate(),112)"
			case "pagenyesterday"
				sql = sql&" and CONVERT(varchar(100),[lasttime],112) = CONVERT(varchar(100),DateAdd(dd,-1,Getdate()),112)"
			case "pagenthismonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate()),0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate())+1,0)))"
			case "pagenlastmonth"
				sql = sql&" and (lasttime between DATEADD(mm,DATEDIFF(mm,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(mm,DATEDIFF(mm,0,getdate()),0)))"
			case "pagenthisyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate()),0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate())+1,0)))"
			case "pagenlastyear"
				sql = sql&" and (lasttime between DATEADD(yy,DATEDIFF(yy,0,Getdate())-1,0) and dateadd(ms,-3,DATEADD(yy,DATEDIFF(yy,0,getdate()),0)))"
	    end select
		if pagendate1 <> "" then sql = sql&" and lasttime >= '"&pagendate1&"'"
		if pagendate2 <> "" then sql = sql&" and lasttime <= '"&pagendate2&"'"
		sql = sql&" group by vpage order by count(id) DESC"
		set Trs = server.createobject("adodb.recordset")
        Trs.Open sql,conn,1,1
		if Trs.eof then msg = "<font color=red>没数据</font>"
		i = 1
        Do While Not Trs.eof and i < 51
			picwidth = Trs("allpage")
			if i = 1 then maxpicwidth = picwidth 
			if maxpicwidth < 600 then
				picwidth = Trs("allpage")
			elseif maxpicwidth >= 600 and maxpicwidth <= 6000 then
				picwidth = int(picwidth/10)
			elseif maxpicwidth >= 6000 and maxpicwidth <= 60000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 60000 and maxpicwidth <= 600000 then
				picwidth = int(picwidth/100)
			elseif maxpicwidth >= 600000 and maxpicwidth <= 6000000 then
				picwidth = int(picwidth/1000)
			else
				picwidth = int(picwidth/10000)
			end if      
		%>
        <tr>
          <td style="width:30%;height:25px;word-wrap:break-word;word-break:break-all;text-align:right;"><a href="<%=Trs("vpage")%>" target="_blank" ><%=left(Trs("vpage"),38)&"..."%></a></td>
          <td style="BORDER-left: #D3D3D3 1px solid;text-align:left;padding-left:0px;"><img src="images/state.jpg" width="<%=picwidth%>" height="17" align="absmiddle">&nbsp;<%=Trs("allpage")%></td>
        </tr>
		<% 
			Trs.movenext
			i = i + 1
        Loop
        Trs.close
		set Trs = nothing
		%>
        <tr>
          <td align="right"><%=msg%></td>
          <td style="BORDER-left: #D3D3D3 1px solid;BORDER-bottom: #D3D3D3 1px solid;"><img src="images/tu.gif" width="1" height="1"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<%
Function pagelist1(currentpage,maxperpage,allpage) 
	currentpage = request.querystring("pageid")
	if isnumeric(currentpage) = false then
		response.write("<script>alert('参数错误！');</script>")
		response.end
	end if
	if currentpage = "" then
		currentpage = 1
	elseif currentpage < 1 then
		currentpage = 1
	else
		currentpage = clng(currentpage)
		if currentpage > allpage then currentpage = allpage			  
	end if
	if not Trs.eof then Trs.move (currentpage-1)*maxperpage
end Function 
Function pagelist2(currentpage,allpage,totalput,sql,book_mark)
%> 
  <table cellpadding="0" cellspacing="0" style="border:none; width:100%;">
  <form name="form1" method="post" action="">
	<tr>
      <td height="20"><p align="center">页数：<%=currentpage%>/<%=allpage%>
              <%
	k=currentpage                                                                                         
   	if k<>1 then
%>
              <a href="?pageid=1&<%=sql%>#<%=book_mark%>">首页</a> <a href="?pageid=<%=k-1%>&<%=sql%>#<%=book_mark%>">上一页</a>
   <%else%>
        首页&nbsp;上一页
   <%end if%>
   <%if k<>allpage then%>
        <a href="?pageid=<%=k+1%>&<%=sql%>#<%=book_mark%>">下一页</a> <a href="?pageid=<%=allpage%>&fromuser=<%=fromuser%>&<%=sql%>#<%=book_mark%>">尾页</a>
        <%else%>
        下一页&nbsp;尾页
        <%end if%>
        共有 <%=totalput%> 条记录
        <input name="pageid" type="text" id="pageid" size="3" class="input1">
        <input type="button" name="Submit" value="GO" onclick=window.location.href="?pageid="+this.form.pageid.value+"&<%=sql%>"; class="bt">
	  </td>
    </tr>
	</form>
  </table>
<%end Function%>
</body>                                                                        
</html>  
<%
function table2(total,table_x,table_y,all_width,all_height,line_no)
	line_color = "#69f"
	left_width = 30
	total_no = ubound(total,1)
	temp1 = 0
	if total_no > 0 then temp6 = total(1,1)
	for i = 1 to total_no
		for j = 1 to line_no
			if temp1 < total(i,j) then temp1 = total(i,j)
			if temp6 > total(i,j) then temp6 = total(i,j)
		next
	next
	temp1 = int(temp1)
	if temp6 > 0 then
		temp6 = int(temp6)
		if temp6 > 10 then
			temp2 = mid(cstr(temp6),2,1)
			if temp2 > 4 then 
				temp3 = (int(temp6/(10^(len(cstr(temp6))-1)))-1)*10^(len(cstr(temp6))-1)
			else
				temp3 = (int(temp6/(10^(len(cstr(temp6))-1)))-0.5)*10^(len(cstr(temp6))-1)
			end if
			temp6 = temp3
		else
			temp6 = 0
		end if
	else
		temp6 = int(0 - temp6)
		if temp6 > 10 then
			temp2 = mid(cstr(temp6),2,1)
			if temp2 > 4 then 
				temp3 = (int(temp6/(10^(len(cstr(temp6))-1)))+1)*10^(len(cstr(temp6))-1)
			else
				temp3 = (int(temp6/(10^(len(cstr(temp6))-1)))+0.5)*10^(len(cstr(temp6))-1)
			end if
			temp6 = 0 - temp3
		else
			temp6 = -10
		end if
	end if
	if temp1 > 9 then
		temp2 = mid(cstr(temp1),2,1)
		if temp2 > 4 then 
			temp3 = (int(temp1/(10^(len(cstr(temp1))-1)))+1)*10^(len(cstr(temp1))-1)
		else
			temp3 = (int(temp1/(10^(len(cstr(temp1))-1)))+0.5)*10^(len(cstr(temp1))-1)
		end if
	else
		if temp1 > 4 then temp3 = 10 else temp3 = 5
	end if
	temp4 = temp3
	Response.write("<v:rect id='_x0000_s1027' alt='' style='position:absolute;left:"&table_x+left_width&"px;top:"&table_y&"px;width:"&all_width&"px;height:"&all_height&"px;z-index:1' fillcolor='#9cf' stroked='f'><v:fill rotate='t' angle='-45' focus='100%' type='gradient'/></v:rect>")
	for i = 0 to all_height step all_height / 5
		Response.write("<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:1' from='"&table_x+left_width+length&"px,"&table_y+all_height-length-i&"px' to='"&table_x+all_width+left_width&"px,"&table_y+all_height-length-i&"px' strokecolor='"&line_color&"'/>")
		Response.write("<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:1' from='"&table_x+(left_width-15)&"px,"&table_y+i&"px' to='"&table_x+left_width&"px,"&table_y+i&"px'/>")

		Response.write("<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x&"px;top:"&table_y+i&"px;width:"&left_width&"px;height:18px;z-index:2'>")
		Response.write("<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='right' style='height:14px;text-align:left;background:none;'>"&temp4&"</td></tr></table></v:textbox></v:shape>")
		temp4 = temp4 - (temp3 - temp6) / 5
	next
	Response.write("<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:1' from='"&table_x+left_width&"px,"&table_y+all_height&"px' to='"&table_x+all_width+left_width&"px,"&table_y+all_height&"px'><v:stroke endarrow='block'/></v:line>")
	Response.write("<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:1' from='"&table_x+left_width&"px,"&table_y&"px' to='"&table_x+left_width&"px,"&table_y+all_height&"px'><v:stroke endarrow='block'/></v:line>")
	Response.write("<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width-50&"px;top:"&table_y-20&"px;width:100px;height:18px;z-index:2'>")
	Response.write("<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='center' style='height:14px;text-align:center;background:none;'>纵坐标</td></tr></table></v:textbox></v:shape>")
	Response.write("<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width+all_width&"px;top:"&table_y+all_height-9&"px;width:100px;height:18px;z-index:2'>")
	Response.write("<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='left' style='height:14px;text-align:left;background:none;'>横坐标</td></tr></table></v:textbox></v:shape>")
	dim line_code
	redim line_code(line_no,5)
	for i = 1 to line_no
		line_temp = split(total(0,i),",")
		line_code(i,1) = line_temp(0)
		line_code(i,2) = line_temp(1)
		line_code(i,3) = line_temp(2)
		line_code(i,4) = line_temp(3)
		line_code(i,5) = line_temp(4)
	next
	for j = 1 to line_no
		for i = 1 to total_no - 1
			x1 = table_x+left_width+all_width*(i-1)/total_no
			y1 = table_y+(temp3-total(i,j))*(all_height/(temp3-temp6))
			x2 = table_x+left_width+all_width*i/total_no
			y2 = table_y+(temp3-total(i+1,j))*(all_height/(temp3-temp6))
			Response.write("<v:line id=""_x0000_s1025"" alt="""" style='position:absolute;left:0;text-align:left;top:0;z-index:2' from="""&x1&"px,"&y1&"px"" to="""&x2&"px,"&y2&"px"" coordsize=""21600,21600"" strokecolor="""&line_code(j,1)&""" strokeweight="""&line_code(j,2)&""">")
			select case line_code(j,3)
				case 1
				case 2
					Response.write("<v:stroke dashstyle='1 1'/>")
				case 3
					Response.write("<v:stroke dashstyle='dash'/>")
				case 4
					Response.write("<v:stroke dashstyle='dashDot'/>")
				case 5
					Response.write("<v:stroke dashstyle='longDash'/>")
				case 6
					Response.write("<v:stroke dashstyle='longDashDot'/>")
				case 7
					Response.write("<v:stroke dashstyle='longDashDotDot'/>")
				case else
			end select
			Response.write("</v:line>"&CHR(13))
			select case line_code(j,4)
				case 1
				case 2
					Response.write("<v:rect id=""_x0000_s1027"" style='position:absolute;left:"&x1-2&"px;top:"&y1-2&"px;width:4px;height:4px; z-index:3' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13))
				case 3
					Response.write("<v:oval id=""_x0000_s1026"" style='position:absolute;left:"&x1-2&"px;top:"&y1-2&"px;width:4px;height:4px;z-index:2' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13))
			end select
			Response.write("<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&x1&"px;top:"&y1-15&"px;width:60px;height:18px;z-index:2'><v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='left' style='height:14px;text-align:left;background:none;'>"&total(i,j)&"</td></tr></table></v:textbox></v:shape>")
		next
		if line_no = 1 then
			x2 = table_x+left_width+all_width*(i-1)/total_no
			y2 = table_y+(temp3-total(i,j))*(all_height/temp3)
		end if 
		select case line_code(j,4)
			case 1
			case 2
				Response.write("<v:rect id=""_x0000_s1027"" style='position:absolute;left:"&x2-2&"px;top:"&y2-2&"px;width:4px;height:4px; z-index:3' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13))
			case 3
				Response.write("<v:oval id=""_x0000_s1026"" style='position:absolute;left:"&x2-2&"px;top:"&y2-2&"px;width:4px;height:4px;z-index:2' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13))
		end select
		Response.write("<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&x2&"px;top:"&y2-15&"px;width:60px;height:18px;z-index:2'><v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='left' style='height:14px;text-align:left;background:none;'>"&total(i,j)&"</td></tr></table></v:textbox></v:shape>")
	next
	for i = 1 to total_no
		Response.write("<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:1' from='"&table_x+left_width+all_width*(i-1)/total_no&"px,"&table_y+all_height&"px' to='"&table_x+left_width+all_width*(i-1)/total_no&"px,"&table_y+all_height+15&"px'/>")

		Response.write("<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width+all_width*(i-1)/total_no&"px;top:"&table_y+all_height&"px;width:"&all_width/total_no&"px;height:18px;z-index:2'>")
		Response.write("<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='left' style='height:14px;text-align:left;background:none;'>"&total(i,0)&"</td></tr></table></v:textbox></v:shape>")
	next
end function
%>
<%call CloseConn()%>