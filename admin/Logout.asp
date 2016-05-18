<!--#include file="../inc/conn.asp"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</HEAD>
<BODY>
<%
if Request.Cookies("admin")("id")<>"" and isnumeric(Request.Cookies("admin")("id")) then
Conn.Execute("Update admin Set sj_no='"&now()&"' Where id=" &Cint(Request.Cookies("admin")("id")))
end if
if not response.cookies("admin").haskeys then
  response.cookies("admin")=""
else
  for each key in response.cookies("admin")
    response.cookies("admin")(key)=""
  next
end if
Response.Cookies("adminpower")=""
response.redirect("Login.asp")
%>
<%call CloseConn()%>
</body>
</html>