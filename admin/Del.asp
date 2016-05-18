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
<div class="ajaxdel">您真的要删除选择的数据吗？<br />删除后不可恢复！</div>
<div class="delsb"><input type="submit" class="bt" value=" 确 定 " /></div>
</form>
<%call CloseConn()%>
</body>
</html>