<!--#include file="chk.asp"-->
<%
mssql = request.QueryString("mssql")
mscode = request.QueryString("mscode")
id = go_num(request.QueryString("id"),0)
Oper = request.QueryString("Oper")
if Oper = "go_off" then Conn.Execute("Update ["&mssql&"] Set "&mscode&" = 0 Where id = "&id)
if Oper = "go_on" then Conn.Execute("Update ["&mssql&"] Set "&mscode&" = 1 Where id = "&id)
Response.Write("<script type = 'text/javascript'>window.parent.location.reload();</script>")
Response.End()
%>
<%call CloseConn()%>
</body>
</html>