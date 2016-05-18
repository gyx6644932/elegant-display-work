<!--#include file="Chk.asp"-->
<%
dim id,title,px,Oper
if Request.form("Oper") <> "" then
	id = go_num(request.form("id"),0)
	title = SafeChar(request.form("title"))
	IF strLength(title) <= 0 then jsouterr("名称不能为空！")
	set rs = server.createobject("adodb.recordset")
	sql = "select * from case_bid where id = "&id
	rs.open sql,conn,1,3
	if Request.form("Oper") = "addsave" then rs.addnew
	rs("title") = title
	rs("px") = go_num(request.form("px"),99)
	rs.update
	rs.close
	set rs = nothing
	Response.Write("<script type='text/javascript'>window.parent.location.reload();</script>")
	response.End()
end if
'传入id，存在则修改，不存在则新增
id = go_num(request.QueryString("id"),0)
if id <= 0 then
	Oper = "addsave"
	px = 99
	title = ""
else
	Oper = "edit"
	Set rs = server.CreateObject("adodb.recordset")
	sql = "select * from case_bid where id = "&id
	rs.Open sql, conn, 1, 1
	px = rs("px")
	title = rs("title")
	rs.close
	set rs = nothing
end if
%>
<table cellpadding="0" cellspacing="1" class="border">
  <form method="post">
  <input type="hidden" name="Oper" value="<%=Oper%>">
  <input type="hidden" name="id" value="<%=id%>">
  <tr class="tdbg">
    <th style="width:25%">分类名称</th>
    <td><input type="text" name="title" value="<%=title%>" maxlength="50" style="width:235px;" /></td>
  </tr>
  <tr class="tdbg">
    <th>排　　序</th>
	<td><input type="text" name="px" value="<%=px%>" class="nochines" maxlength="4" style="width:60px;" /></td>
  </tr>
  <tr class="tdbg2">
	<td colspan="2"><input type="submit" class="bt" value="确定" /></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>