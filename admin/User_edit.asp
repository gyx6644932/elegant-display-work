<!--#include file="chk.asp"-->
<%
id = Request.QueryString("id")
if Request.Form("Oper") = "edit" then
	Response.Expires = 0
	id = Request.Form("id")
	realname = replace(trim(Request("name")),"'","")
	if realname = "" or isnull(realname) then
		Response.write("<script>alert('������������Ϊ�գ�');location.href='Javascript:history.back()';</script>")
		Response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from admin where id = "&id
	rs.open sql,conn,3,3
	rs("name") = realname
	rs.update
	rs.close
	set rs = nothing
	Response.write("<script>alert('�޸ĳɹ�����Ҫ���µ�¼����Ч��');location.href='Main.asp';</script>")
	Response.end
end if
set rs = server.createobject("adodb.recordset")
sql = "select * from admin where id = "&id
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
function chk_data(){
	if (document.getElementById("name").value.length <= 0){
		alert("\n�������������ǲ�������Ŀ���ԭ��\n\n����������Ϊ�գ�");
		document.getElementById("name").focus();
		return false;
	}
}
</script>
<table cellpadding="0" cellspacing="1" class="border">
  <form method="post" onSubmit="JavaScript: return chk_data();">
  <input type="hidden" name="id" value="<%=rs("id")%>">
  <input type="hidden" name="Oper" value="edit">
  <tr><th colspan="2">�޸ĸ�������</th></tr>
  <tr class="tdbg">
    <td class="t-right h25">�� �� ����</td>
    <td><input type="text" value="<%=rs("Userid")%>" class="input-text" disabled></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�����ݣ�</td>
    <td><input type="text" value="<%=rs("title")%>" size="30" class="input-text" disabled></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�ա�������</td>
    <td><input name="name" type="text" value="<%=rs("name")%>" size="30" maxlength="30" class="input-text"> *</td>
  </tr>
  <tr class="tdbg2">
    <td colspan="2"><input type="submit" class="bt" value="ȷ���ύ">����<input type="reset" class="bt" value="������д"></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>