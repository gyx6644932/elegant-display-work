<!--#include file="chk.asp"-->
<%
if Request.Form("Oper") = "addsave" then
	Response.Expires = 0
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
	Userid = ReplaceBadChar(Request.Form("Userid"))
	if Userid = "" or isnull(Userid) or strLength(Userid) < 5 then
		Response.write("<script>alert('�����û�������Ϊ�գ��Ҳ���С��5λ��');location.href='Javascript:history.back()';</script>")
		Response.end
	end if
	title = ReplaceBadChar(Request.Form("title"))
	if title = "" or isnull(title) then
		Response.write("<script>alert('������ݲ���Ϊ�գ�');location.href='Javascript:history.back()';</script>")
		Response.end
	end if
	pass = Request.Form("pass")
	adminpower = ""
	for i=0 to 99
		adminpower = adminpower&"0|"
	next
	adminpower = adminpower&"0"
	'�������ݿ⣬���˹���Ա�Ƿ��Ѿ�����
	set rs = server.createobject("adodb.recordset")
	sql="select * from admin where Userid = '"&Userid&"'"
	rs.open sql,conn,1,1
	if rs.recordcount >= 1 then
		response.write("<script language=javascript>alert('�˹���Ա�ʺ��Ѿ����ڣ���ѡ�������ʺ�!');history.go(-1);</script>")
		response.End
		rs.close
		set rs = nothing
	end if
	set rs = server.createobject("adodb.recordset")
	sql = "select * from admin where id = 0"
	rs.open sql,conn,3,3
	'���һ������Ա�ʺŵ����ݿ�
	rs.addnew
	rs("Userid") = Userid
	rs("userpsw") = md5(userpsw)
	rs("title") = title
	rs("name") = names
	rs("pass") = pass
	rs("userkey") = 0
	rs("gonum") = 0
	rs("goip") = 0
	rs("lasttime") = "���� 12:00:00"
	rs("sj_no") = strj_no(now())
	rs("adminpower") = adminpower
	rs.update
	rs.close
	set rs = nothing
	Response.Redirect("Admin.asp")
	Response.End()
end if
%>
<script language = "JavaScript">
function chk_data(){
	if (document.getElementById("userid").value.length < 5){
		alert("\n�������������ǲ�������Ŀ���ԭ��\n\n���û�������С��5λ��");
		document.getElementById("userid").focus();
		return false;
	}
	if (document.getElementById("title").value.length <= 0){
		alert("\n�������������ǲ�������Ŀ���ԭ��\n\n���û���ݲ���Ϊ�գ�");
		document.getElementById("title").focus();
		return false;
	}
	if (document.getElementById("name").value.length <= 0){
		alert("\n�������������ǲ�������Ŀ���ԭ��\n\n���û���������Ϊ�գ�");
		document.getElementById("name").focus();
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
  <input type="hidden" name="Oper" value="addsave">
  <tr><th colspan="2">��������Ա</th></tr>
  <tr class="tdbg">
    <td class="t-right h25">�� �� ����</td>
    <td><input name="userid" type="text" id="userid" size="20" maxlength="20" class="input-text"> 5-20���ַ�</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�����ݣ�</td>
    <td><input name="title" type="text" size="30" maxlength="8" class="input-text"> �磺����Ա���ܹ���Ա������רԱ �ȵ��Զ���ƺ�</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�ա�������</td>
    <td><input name="name" type="text" size="30" maxlength="8" class="input-text"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�� �� �룺</td>
    <td><input name="newpsw" type="password" id="newpsw" size="40" maxlength="30" class="input-text"> 6-30���ַ�</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">ȷ�����룺</td>
    <td><input name="endpsw" type="password" id="endpsw" size="40" maxlength="30" class="input-text"> 6-30���ַ�</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">״̬���ƣ�</td>
    <td><input name="pass" type="radio" value="1" checked>��ͨ����<input name="pass" type="radio" value="0">�ر�</td>
  </tr>
  <tr class="tdbg2">
    <td colspan="2"><input type="submit" class="bt" value="ȷ���ύ">����<input type="reset" class="bt" value="������д"></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>