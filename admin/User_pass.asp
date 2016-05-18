<!--#include file="chk.asp"-->
<%
id = Request.QueryString("id")
if Request.Form("Oper") = "edit" then
	Response.Expires = 0
	id = Request.Form("id")
	userpsw = replace(trim(Request("newpsw")),"'","")
	if userpsw = "" or isnull(userpsw) then
		Response.write("<script>alert('出错：密码不能为空！');location.href='Javascript:history.back()';</script>")
		Response.end
	end if
	set rs = server.createobject("adodb.recordset")
	sql = "select * from admin where id = "&id
	rs.open sql,conn,1,3
	rs("userpsw") = md5(userpsw)
	rs.update
	rs.close
	set rs = nothing
	Response.Redirect("Main.asp")
	Response.End()
end if
set rs = server.createobject("adodb.recordset")
sql = "select * from admin where id = "&id
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
function chk_data(){
	if (document.getElementById("newpsw").value.length < 6){
		alert("\n操作出错，下面是产生错误的可能原因：\n\n·密码不能小于6位！");
		document.getElementById("newpsw").focus();
		return false;
	}
	if ((document.getElementById("newpsw").value)!=(document.getElementById("endpsw").value)){
		alert("\n操作出错，下面是产生错误的可能原因：\n\n·确认密码和登录密码不一致，请重新输入!");
		document.getElementById("endpsw").focus();
		return false;
	}
}
</script>
<table cellpadding="0" cellspacing="1" class="border">
  <form method="post" onSubmit="JavaScript: return chk_data();">
  <input type="hidden" name="Oper" value="edit">
  <input type="hidden" name="id" value="<%=rs("id")%>">
  <tr><th colspan="2">修改登录密码</th></tr>
  <tr class="tdbg">
    <td class="t-right h25">用 户 名：</td>
    <td><input type="text" value="<%=rs("Userid")%>" class="input-text" disabled></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">新 密 码：</td>
    <td><input type="password" name="newpsw" id="newpsw" size="50" maxlength="30" class="input-text"> 6-30个字符</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">确认密码：</td>
    <td><input type="password" name="endpsw" id="endpsw" size="50" maxlength="30" class="input-text"> 6-30个字符</td>
  </tr>
  <tr class="tdbg2">
    <td colspan="2"><input type="submit" class="bt" value="确定提交">　　<input type="reset" class="bt" value="重新填写"></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>