<!--#include file="chk.asp"-->
<%
if Request.form("Oper") <> "" then
	title = ReplaceBadChar(request.form("title"))
	url = request.form("url")
	id = ReplaceBadChar(request.form("id"))
	px = ReplaceBadChar(request.form("px"))
	IF strLength(id) <= 0 or not isshuzi(id) then jsouterr("�����������")
	IF strLength(px) <= 0 or not isshuzi(px) then jsouterr("�������Ϊ���֣��Ҳ���Ϊ�գ�")
	IF strLength(url) <= 0 then jsouterr("���Ӳ���Ϊ�գ�")
	IF strLength(title) > 20 then jsouterr("��վ����̫����������һЩ��")
	set rs = server.createobject("adodb.recordset")
	sql = "select * from link where id = "&id
	rs.open sql,conn,1,3
	if Request.form("Oper") = "addsave" then rs.addnew
	rs("url") = url
	rs("title") = title
	rs("px") = px
	rs.update
	rs.close
	set rs = nothing
	Response.Write("<script type='text/javascript'>window.parent.location.reload();</script>")
	response.End()
end if
'����id���������޸ģ�������������
id = request.QueryString("id")
if id <> "" and isnumeric(id) then
	Oper = "edit"
	Set rs = server.CreateObject("adodb.recordset")
	sql = "select * from link where id = "&id
	rs.Open sql, conn, 1, 1
	px = rs("px")
	title = rs("title")
	url = rs("url")
	rs.close
	set rs = nothing
else
	Oper = "addsave"
	title = ""
	url = "http://"
	px = 99
	id = 0
end if
%>
<table cellpadding="0" cellspacing="1" class="border">
  <form method="post">
  <input type="hidden" name="Oper" value="<%=Oper%>">
  <input type="hidden" name="id" value="<%=id%>">
  <tr class="tdbg">
    <th class="w30">�š�����</th>
	<td><input type="text" name="px" value="<%=px%>" class="w100 nochines" maxlength="8" /></td>
  </tr>
  <tr class="tdbg">
    <th>��վ����</th>
    <td><input type="text" name="title" value="<%=title%>" class="w100" maxlength="45" /></td>
  </tr>
  <tr class="tdbg">
    <th>��վ��ַ</th>
    <td><input name="url" type="text" value="<%=url%>" class="w100" maxlength="240" /></td>
  </tr>
  <tr class="tdbg2">
	<td colspan="2"><input type="submit" class="bt" value="ȷ��" /></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>