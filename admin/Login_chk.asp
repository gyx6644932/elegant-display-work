<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/md5.asp"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</HEAD>
<BODY>
<%
if request("cookieexists")=false then
response.Write "<br><p align=center><font color='red' size='9pt'>���棡�����������Java�ǹر�״̬��������ͼ�Ƿ���¼�����������ٵ�¼��</font></p>"
response.end
end if
username=replace(trim(request("username")),"'","")
userpsw=md5(replace(trim(request("userpsw")),"'",""))
checkcode=replace(trim(request("checkcode")),"'","")
if cstr(session("getcode"))<>cstr(trim(request("checkcode"))) then
response.Write "<script LANGUAGE='javascript'>alert('��������ȷ����֤�룡');history.go(-1);</script>"
response.end
response.Redirect"login.asp"
end if
set rs=server.createobject("adodb.recordset")
sql="select * from Admin where userid='"&username&"' and userpsw='"&userpsw&"' and pass=1"
rs.open sql,conn,1,3
if rs.bof and rs.eof then
response.write "<script LANGUAGE='javascript'>alert('�Բ��������û������������� ����������û����������ˣ�');history.go(-1);</script>"
response.End()
else
sj_no=strj_no(now())
rs("sj_no")=sj_no
if rs("gonum")<> 0 then
	Response.Cookies("admin")("lasttime")=rs("lasttime")
else
	Response.Cookies("admin")("lasttime")="��һ�ε�¼"
end if
rs("gonum")=rs("gonum")+1
Response.Cookies("admin")("oldip")=rs("goip")
rs("goip")=request.ServerVariables("REMOTE_HOST")
rs("lasttime")=now()
rs.update
Response.Cookies("admin")("id")=rs("id")
Response.Cookies("admin")("user")=username
Response.Cookies("admin")("psw")=md5(md5(sj_no)&userpsw)
Response.Cookies("admin")("gonum")=rs("gonum")
if rs("name")<>"" then
Response.Cookies("admin")("name")=rs("name")
else
Response.Cookies("admin")("name")="δ֪"
end if
if rs("title")<>"" then
Response.Cookies("admin")("title")=rs("title")
else
Response.Cookies("admin")("title")="δ֪"
end if
Response.Cookies("admin")("ip")=rs("goip")
Response.Cookies("adminpower")=rs("adminpower")
rs.close
set rs=nothing
response.Redirect"index.asp"
end if
%>
<%call CloseConn()%>
</body>
</html>