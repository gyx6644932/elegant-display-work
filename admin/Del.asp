<!--#include file="chk.asp"-->
<style type="text/css">*,body { margin:0; padding:0; overflow:hidden;}</style>
<%
if trim(Request.form("Oper")) = "del" then
	id = go_num(request.form("id"),0)
	mssql = trim(Request.form("mssql"))
	Conn.Execute("delete from ["&mssql&"] where id = "&id)
	Call CloseConn()
	Response.Write("<script type='text/javascript'>window.parent.location.reload();</script>")
	response.End()
end if
%>
<form method="post">
<input type="hidden" name="id" value="<%=go_num(request.QueryString("id"),0)%>" />
<input type="hidden" name="Oper" value="del" />
<input type="hidden" name="mssql" value="<%=trim(Request.QueryString("mssql"))%>">
<div class="ajaxdel">�����Ҫɾ��ѡ���������<br />ɾ���󲻿ɻָ���</div>
<div class="delsb"><input type="submit" class="bt" value=" ȷ �� " /></div>
</form>
<%call CloseConn()%>
</body>
</html>