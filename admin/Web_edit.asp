<!--#include file="chk.asp"-->
<%
if Request.form("Oper") <> "" then
	d1 = ReplaceBadChar(request.form("d1"))
	d2 = ReplaceBadChar(request.form("d2"))
	title = ReplaceBadChar(request.form("title"))
	content = request.form("content")
	id = request.form("id")
	bid = request.form("bid")
	IF strLength(id) <= 0 or not isshuzi(id) then jsouterr("�����������")
	set rs = server.createobject("adodb.recordset")
	sql = "select * from web where id = "&id
	rs.open sql,conn,1,3
	if Request.form("Oper") = "addsave" then rs.addnew
	rs("d1") = d1
	rs("d2") = d2
	rs("title") = title
	rs("content") = content
	rs("bid") = bid
	rs.update
	rs.close
	set rs = nothing
	call jsoutgo("�����ɹ���","Web.asp")
end if
'����id���������޸ģ�������������
id = request.QueryString("id")
if id <> "" and isnumeric(id) then
	Oper = "edit"
	Set rs = server.CreateObject("adodb.recordset")
	sql = "select * from web where id = "&id
	rs.Open sql, conn, 1, 1
	title = rs("title")
	content = rs("content")
	bid = rs("bid")
	d1 = rs("d1")
	d2 = rs("d2")
	rs.close
	set rs = nothing
else
	Oper = "addsave"
	d1 = webc("d1")
	d2 = webc("d2")
	bid = 0
	id = 0
end if
%>
<script charset="utf-8" src="../html/kindeditor-min.js"></script>
<script charset="utf-8" src="../html/lang/zh_CN.js"></script>
<script>
var editor;
KindEditor.ready(function(K){
	editor = K.create('#content',{allowFileManager : true});
});
</script>
<!--#include file="Web_top.asp"-->
<table cellpadding="2" cellspacing="1"class="border">
  <form method="post">
  <input type="hidden" name="Oper" value="<%=Oper%>">
  <input type="hidden" name="id" value="<%=id%>">
  <tr><th colspan="2">�޸ĵ�ҳ</th></tr>
  <tr class="tdbg">
    <td class="t-right h25">��ҳģʽ</td>
    <td><input name="bid" type="radio" value="0" <%if bid = 0 then response.write("checked")%>>��ͳ���֡�<input name="bid" type="radio" value="1" <%if bid = 1 then response.write("checked")%>>ȫ����ҳ</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">��ҳ����</td>
    <td><input type="text" name="title" value="<%=title%>" class="w100" maxlength="45"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�� �� ��</td>
    <td><input type="text" name="d1" value="<%=d1%>" class="w100" maxlength="240"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">��������</td>
    <td><input type="text" name="d2" value="<%=d2%>" class="w100" maxlength="240"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">ҳ������</td>
    <td><textarea name="content" id="content" style="width:100%;height:350px;"><%=content%></textarea></td>
  </tr>
  <tr class="tdbg2">
    <td colspan="2"><input type="submit" class="bt" value="ȷ���ύ">����<input type="reset" class="bt" value="������д"></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>