<!--#include file="chk.asp"-->
<%
'****************************************�ύ*********************************************
if Request.form("Oper") <> "" then
	id = ReplaceBadChar(request.form("id"))
	bid = split(request.form("bid"),"||")
	px = ReplaceBadChar(request.form("px"))
	IF strLength(id) <= 0 or not isshuzi(id) then jsouterr("�����������")
	IF strLength(px) <= 0 or not isshuzi(px) then jsouterr("�������Ϊ���֣��Ҳ���Ϊ�գ�")
	set rs = server.createobject("adodb.recordset")
	sql = "select * from news where id = "&id
	rs.open sql,conn,1,3
	if Request.form("Oper") = "addsave" then rs.addnew
	rs("title") = request.form("title")
	rs("bid") = bid(0)
	rs("bname") = bid(1)
	rs("content") = request.form("content")
	rs("op_come") = request.form("op_come")
	rs("addtime") = now()
	rs("px") = px
	rs.update
	rs.close
	set rs = nothing
	call jsoutgo("�����ɹ���","news.asp")
end if
'****************************************����ӳ�ʼ��*********************************************
'����id���������޸ģ�������������
id = request.QueryString("id")
if id <> "" and isnumeric(id) then
	Oper = "edit"
	Set rs = server.CreateObject("adodb.recordset")
	sql = "select * from news where id = "&id
	rs.Open sql, conn, 1, 1
	title = rs("title")
	bid = rs("bid")
	content = rs("content")
	op_come = rs("op_come")
	px = rs("px")
	rs.close
	set rs = nothing
else
	Oper = "addsave"
	bid = 0
	op_come = "��վ"
	bx2 = 5
	px = 99
	id = 0
end if
%>
<script type="text/javascript" src="../js/colorselect.js"></script>
<script charset="utf-8" src="../html/kindeditor-min.js"></script>
<script charset="utf-8" src="../html/lang/zh_CN.js"></script>
<script type="text/javascript">
var editor;
KindEditor.ready(function(K){
	editor = K.create('#content',{allowFileManager : true});
});
</script>
<table cellpadding="0" cellspacing="1" class="border">
  <tr><th colspan="2">���ű༭</th></tr>
  <form method="post">
  <input name="id" type="hidden" value="<%=id%>" />
  <input type="hidden" name="Oper" value="<%=Oper%>" />
  <tr class="tdbg">
    <td class="t-right h25 w10">���ű���</td>
    <td><input type="text" name="title" value="<%=title%>" class="w100" maxlength="240" /></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">���ŷ���</td>
    <td><select name="bid" id="bid">
<%
set rsb = server.CreateObject("adodb.recordset")
sql = "Select * from news_bid order by px asc"
rsb.Open sql,Conn,1,1
if not rsb.eof then
if bid = 0 then bid = rsb("id")
do while not rsb.eof
if Cstr(rsb("id")) = Cstr(bid) then bid = rsb("id")
%>
<option value="<%=rsb("id")%>||<%=rsb("title")%>" <%if Cstr(rsb("id")) = Cstr(bid) then Response.Write("selected")%>><%=rsb("title")%></option>
<%
rsb.movenext
loop
end if
rsb.close
set rsb=nothing
%></select></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right">��������</td>
    <td><textarea name="content" id="content" style="width:100%;height:350px;"><%=content%></textarea></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right">������Դ</td>
    <td><input type="text" name="op_come" value="<%=op_come%>" class="w100" maxlength="45" /></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right">�š�����</td>
    <td><input type="text" name="px" value="<%=px%>" class="w10 nochines" maxlength="8" /></td>
  </tr>
  <tr class="tdbg2">
    <td colspan="2"><input type="submit" class="bt" value="ȷ���ύ" /></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>