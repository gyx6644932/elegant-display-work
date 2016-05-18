<!--#include file="chk.asp"-->
<%
if Request.form("Oper") <> "" then
	title = ReplaceBadChar(request.form("title"))
	content = request.form("content")
	id = ReplaceBadChar(request.form("id"))
	px = ReplaceBadChar(request.form("px"))
	IF strLength(id) <= 0 or not isshuzi(id) then jsouterr("发生意外错误！")
	IF strLength(px) <= 0 or not isshuzi(px) then jsouterr("排序必须为数字，且不能为空！")
	IF strLength(content) <= 0 then jsouterr("链接不能为空！")
	IF strLength(title) > 36 then jsouterr("条目名称太长！")
	set rs = server.createobject("adodb.recordset")
	sql = "select * from xt_quick where id = "&id
	rs.open sql,conn,1,3
	if Request.form("Oper") = "addsave" then rs.addnew
	rs("content") = content
	rs("title") = title
	rs("px") = px
	rs.update
	rs.close
	set rs = nothing
	Response.Write("<script type='text/javascript'>window.parent.location.reload();</script>")
	response.End()
end if
'传入id，存在则修改，不存在则新增
id = request.QueryString("id")
if id <> "" and isnumeric(id) then
	Oper = "edit"
	Set rs = server.CreateObject("adodb.recordset")
	sql = "select * from xt_quick where id = "&id
	rs.Open sql, conn, 1, 1
	px = rs("px")
	title = rs("title")
	content = rs("content")
	rs.close
	set rs = nothing
else
	Oper = "addsave"
	title = ""
	content = ""
	px = 99
	id = 0
end if
%>
<table cellpadding="0" cellspacing="1" class="border">
  <form method="post">
  <input type="hidden" name="Oper" value="<%=Oper%>">
  <input type="hidden" name="id" value="<%=id%>">
  <tr class="tdbg">
    <th class="w30">排　　序</th>
	<td><input type="text" name="px" value="<%=px%>" class="w100 nochines" maxlength="8" /></td>
  </tr>
  <tr class="tdbg">
    <th>条目名称</th>
    <td><input type="text" name="title" value="<%=title%>" class="w100" maxlength="45" /></td>
  </tr>
  <tr class="tdbg">
    <th>链接地址</th>
    <td><input name="content" type="text" value="<%=content%>" class="w100 nochines" maxlength="240" /></td>
  </tr>
  <tr class="tdbg2">
	<td colspan="2"><input type="submit" class="bt" value="确定" /></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>