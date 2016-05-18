<!--#include file="chk.asp"-->
<%
mssql = request.QueryString("mssql")
id = request.QueryString("id")
if request("Oper") = "edit" then
	px = request.form("px")
	id = request.form("id")
	mssql = request.Form("mssql")
	if px <> "" and isNumeric(px) then
		set rs = server.createobject("adodb.recordset")
		sql = "select id,px from ["&mssql&"] where id = "&id
		rs.open sql,conn,1,3
		rs("px") = px
		rs.update
		rs.close
		set rs=nothing
		Response.Write("<script type = 'text/javascript'>window.parent.location.reload();</script>")
	else
		Response.Write("<script type = 'text/javascript'>alert('排序值可以是小于0的负数，但必须为数字！越小越靠前！');window.location.href='Javascript:history.back()';</script>")
	end if
	Response.End()
end if
set rs = server.createobject("adodb.recordset")
sql = "select id,px from ["&mssql&"] where id = "&id
rs.open sql,conn,1,1
%>
<table cellpadding="0" cellspacing="0" class="w100">
  <form name="formadd" method="post">
  <input type="hidden" name="Oper" value="edit" />
  <input type="hidden" name="mssql" value="<%=mssql%>" />
  <input type="hidden" name="id" value="<%=rs("id")%>" />
  <tr>
    <td style="padding-left:30px;"><input type="text" name="px" value="<%=rs("px")%>" style="ime-mode:disabled;width:120px;" maxlength="10" /></td>
	<td style="width:60px;" align="center"><input type="submit" class="bt" value="提交" /></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>