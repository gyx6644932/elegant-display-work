<!--#include file="chk.asp"-->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
bid = go_num(Request.QueryString("bid"),0)
typejj = go_num(Request.QueryString("typejj"),1)
Keyword = ReplaceBadChar(Request.QueryString("Keyword"))
if request("Oper")="del" then
	id = request.Form("id")
	if id <> "" then
		id = split(id,",")
		for i = 0 to UBound(id)
		conn.execute("delete from news where id="&id(i)&"")
		next
	end if
	Response.Redirect("news.asp")
	Response.End()
end if
%>
<SCRIPT type="text/javascript">
function CheckAll(form){
	for (var i=0;i<form.elements.length;i++){
		var e = form.elements[i];
		if (e.name != 'chkall') e.checked = form.chkall.checked; 
	}
}
</SCRIPT>
<table cellpadding="0" cellspacing="1"  class="border">
  <tr><th>新闻管理</th></tr>
  <form method="get">
  <tr class="tdbg">
    <td class="t-left h25"><select size="1" name="bid">
<option value="0">全部新闻</option>
<%
set rs = server.CreateObject("adodb.recordset")
sql = "Select * from news_bid order by px asc,id desc"
rs.Open sql,Conn,1,1
if not rs.eof then
do while not rs.eof
%>
<option value="<%=rs("id")%>" <%if Cstr(rs("id")) = Cstr(bid) then Response.Write("selected")%>><%=rs("title")%></option>
<%
rs.movenext
loop
end if
rs.close
set rs=nothing
%></select> <input type="text" name="Keyword" size="20" value="<%=Keyword%>"> <select size="1" name="typejj">
<option value="1" <%if typejj = 1 then Response.Write("selected")%>>标题内容</option>
<option value="2" <%if typejj = 2 then Response.Write("selected")%>>文章内容</option>
<option value="3" <%if typejj = 3 then Response.Write("selected")%>>新闻来源</option>
</select> <input type="submit" class="bt" value="搜索"></td>
  </tr>
  </form>
</table>
<br>
<table cellpadding="0" cellspacing="1" class="border">
  <form method="Post" name="delform">
  <input type="hidden" name="Oper" value="del">
  <tr>
    <th><input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)"></th>
    <th>排序</th>
    <th>分类</th>
    <th>标题</th>
    <th>来源</th>
    <th>发布时间</th>
    <th>修改</th>
	<th>删除</th>
  </tr> 
<%
set rs = Server.CreateObject("ADODB.Recordset")
sql = "Select * From news Where 1 = 1"
if bid <> 0 then sql = sql&" and bid = "&bid
if Keyword <> "" then
	if typejj = 1 then
		sql = sql&" and title like '%"&Keyword&"%'"
	elseif typejj = 2 then
		sql = sql&" and content like '%"&Keyword&"%'"
	elseif typejj = 3 then
		sql = sql&" and op_come like '%"&Keyword&"%'"
	end if
end if
sql = sql&" Order By px asc,id Desc"
rs.Open sql,Conn,1,1
if not rs.EOF then
	rs.PageSize = 20 
    iCount = rs.RecordCount 
    iPageSize = rs.PageSize
    maxpage = rs.PageCount
    page = request("page")
    If Not IsNumeric(page) Or page = "" Then
        page = 1
    Else
        page = CInt(page)
    End If
    If page<1 Then
        page = 1
    ElseIf page>maxpage Then
        page = maxpage
    End If
    rs.AbsolutePage = Page
    If page = maxpage Then
        x = iCount - (maxpage -1) * iPageSize
    Else
        x = iPageSize
    End If
    For i = 1 To x
%>	  
  <tr class="tdbg" onMouseOver="this.className='tdbg3'" onMouseOut="this.className='tdbg'">
    <td><input type='checkbox' name="ID" value="<%=rs("id")%>"></td>
    <td class="t-center h25"><a href="Desc.asp?id=<%=rs("id")%>&mssql=news&FB_iframe=true&height=50&width=210" class="firstebox" title="修改排序值"><%=rs("px")%></a></td>
    <td class="t-center"><%=rs("bname")%></td>
    <td class="t-left"><a href="../index/newsshow.asp?id=<%=rs("id")%>" target="_blank"><%=gotTopic(rs("title"),44)%></a></td>
    <td class="t-center"><%=rs("op_come")%></td>
    <td class="t-center"><%=FormatTime(rs("addtime"),8)%></td>
    <td class="t-center"><a href="news_edit.asp?id=<%=rs("id")%>"><img src="images/edit/edit.gif"></a></td>
	<td class="t-center"><a href="Del.asp?id=<%=rs("id")%>&mssql=news&FB_iframe=true&height=100&width=200" class="firstebox" title="删除确认"><img src="images/edit/adel.gif"></a></td>
  </tr>
<%
	rs.movenext
	Next
End If
%>
  </form>
    <tr class="tdbg2">
    <td colspan="8">
      <table cellpadding="0" cellspacing="0" class="w100">
        <tr>
          <td><input type="button" onClick="if(confirm('确定要删除选中的信息吗？一旦删除将不能恢复！'))delform.submit()" class="bt" value="删除选中"></td>
          <td class="w100"><%Call PageControl(iCount, maxpage, page, iPageSize)%></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%call CloseConn()%>
</body>
</html>