<!--#include file="chk.asp"-->
<%
id = Request.QueryString("id")
if Request("Oper") = "edit" then
	response.Expires = 0
	id = Request.Form("id")
	userpsw = Request.Form("newpsw")
	if userpsw = "" or isnull(userpsw) or strLength(userpsw) < 6 then
		Response.write("<script>alert('�������벻��Ϊ�գ��Ҳ���С��6λ��');location.href='Javascript:history.back()';</script>")
		Response.end
	end if
	names = ReplaceBadChar(Request.Form("name"))
	if names = "" or isnull(names) then
		Response.write("<script>alert('������������Ϊ�գ�');location.href='Javascript:history.back()';</script>")
		Response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from admin where id = "& id
	rs.open sql,conn,1,3
	rs("userpsw") = md5(userpsw)
	rs("name") = names
	rs("title") = ReplaceBadChar(Request.Form("title"))
	rs.update
	rs.close
	set rs = nothing
	Response.Redirect("Admin.asp")
end if
set rs = server.createobject("adodb.recordset")
sql = "select * from admin where id = "&id
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
function chk_data(){
	if (document.getElementById("name").value.length <= 0){
		alert("\n�������������ǲ�������Ŀ���ԭ��\n\n���û���������Ϊ�գ�");
		document.getElementById("name").focus();
		return false;
	}
	if (document.getElementById("title").value.length <= 0){
		alert("\n�������������ǲ�������Ŀ���ԭ��\n\n���û���ݲ���Ϊ�գ�");
		document.getElementById("title").focus();
		return false;
	}
	if (document.getElementById("newpsw").value.length < 6){
		alert("\n�������������ǲ�������Ŀ���ԭ��\n\n�����벻��С��6λ��");
		document.getElementById("newpsw").focus();
		return false;
	}
	if (document.getElementById("newpsw").value != document.getElementById("endpsw").value){
		alert("\n�������������ǲ�������Ŀ���ԭ��\n\n��ȷ������͵�¼���벻һ�£����������룡");
		document.getElementById("newpsw").focus();
		return false;
	}
}
</script>
<!--#include file="Admin_top.asp"-->
<table cellpadding="0" cellspacing="1" class="border">
  <form method="post" onSubmit="JavaScript: return chk_data();">
  <input type="hidden" name="Oper" value="edit">
  <input type="hidden" name="id" value="<%=rs("id")%>">
  <tr><th colspan="2">�޸����뼰ʹ����</td></tr>
  <tr class="tdbg">
    <td class="t-right h25">�� �� ����</td>
    <td><input type="text" value="<%=rs("Userid")%>" class="input-text" disabled></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�ա�������</td>
    <td><input type="text" name="name" value="<%=rs("name")%>" class="input-text"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�����ݣ�</td>
    <td><input type="text" name="title" value="<%=rs("title")%>" class="input-text"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�� �� �룺</td>
    <td><input type="password" name="newpsw" size="40" maxlength="30" class="input-text"> 6-30���ַ�</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">ȷ�����룺</td>
    <td><input type="password" name="endpsw" size="40" maxlength="30" class="input-text"> 6-30���ַ�</td>
  </tr>
  <tr class="tdbg2">
    <td colspan="2"><input type="submit" class="bt" value="ȷ���ύ">����<input type="reset" class="bt" value="������д"></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>