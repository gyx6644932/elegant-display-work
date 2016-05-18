<!--#include file="chk.asp"-->
<%
if Request.Form("Oper") = "addsave" then
	Response.Expires = 0
	userpsw = Request.Form("newpsw")
	if userpsw = "" or isnull(userpsw) or strLength(userpsw) < 6 then
		Response.write("<script>alert('出错：密码不能为空，且不能小于6位！');location.href='Javascript:history.back()';</script>")
		Response.end
	end if
	names = ReplaceBadChar(Request.Form("name"))
	if names = "" or isnull(names) then
		Response.write("<script>alert('出错：姓名不能为空！');location.href='Javascript:history.back()';</script>")
		Response.end
	end if
	Userid = ReplaceBadChar(Request.Form("Userid"))
	if Userid = "" or isnull(Userid) or strLength(Userid) < 5 then
		Response.write("<script>alert('出错：用户名不能为空，且不能小于5位！');location.href='Javascript:history.back()';</script>")
		Response.end
	end if
	title = ReplaceBadChar(Request.Form("title"))
	if title = "" or isnull(title) then
		Response.write("<script>alert('出错：身份不能为空！');location.href='Javascript:history.back()';</script>")
		Response.end
	end if
	pass = Request.Form("pass")
	adminpower = ""
	for i=0 to 99
		adminpower = adminpower&"0|"
	next
	adminpower = adminpower&"0"
	'查找数据库，检查此管理员是否已经存在
	set rs = server.createobject("adodb.recordset")
	sql="select * from admin where Userid = '"&Userid&"'"
	rs.open sql,conn,1,1
	if rs.recordcount >= 1 then
		response.write("<script language=javascript>alert('此管理员帐号已经存在，请选用其他帐号!');history.go(-1);</script>")
		response.End
		rs.close
		set rs = nothing
	end if
	set rs = server.createobject("adodb.recordset")
	sql = "select * from admin where id = 0"
	rs.open sql,conn,3,3
	'添加一个管理员帐号到数据库
	rs.addnew
	rs("Userid") = Userid
	rs("userpsw") = md5(userpsw)
	rs("title") = title
	rs("name") = names
	rs("pass") = pass
	rs("userkey") = 0
	rs("gonum") = 0
	rs("goip") = 0
	rs("lasttime") = "上午 12:00:00"
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
		alert("\n操作出错，下面是产生错误的可能原因：\n\n·用户名不能小于5位！");
		document.getElementById("userid").focus();
		return false;
	}
	if (document.getElementById("title").value.length <= 0){
		alert("\n操作出错，下面是产生错误的可能原因：\n\n·用户身份不能为空！");
		document.getElementById("title").focus();
		return false;
	}
	if (document.getElementById("name").value.length <= 0){
		alert("\n操作出错，下面是产生错误的可能原因：\n\n·用户姓名不能为空！");
		document.getElementById("name").focus();
		return false;
	}
	if (document.getElementById("newpsw").value.length < 6){
		alert("\n操作出错，下面是产生错误的可能原因：\n\n·密码不能小于6位！");
		document.getElementById("newpsw").focus();
		return false;
	}
	if (document.getElementById("newpsw").value != document.getElementById("endpsw").value){
		alert("\n操作出错，下面是产生错误的可能原因：\n\n·确认密码和登录密码不一致，请重新输入！");
		document.getElementById("newpsw").focus();
		return false;
	}
}
</script>
<!--#include file="Admin_top.asp"-->
<table cellpadding="0" cellspacing="1" class="border">
  <form method="post" onSubmit="JavaScript: return chk_data();">
  <input type="hidden" name="Oper" value="addsave">
  <tr><th colspan="2">新增管理员</th></tr>
  <tr class="tdbg">
    <td class="t-right h25">用 户 名：</td>
    <td><input name="userid" type="text" id="userid" size="20" maxlength="20" class="input-text"> 5-20个字符</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">身　　份：</td>
    <td><input name="title" type="text" size="30" maxlength="8" class="input-text"> 如：管理员、总管理员、刊登专员 等等自定义称呼</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">姓　　名：</td>
    <td><input name="name" type="text" size="30" maxlength="8" class="input-text"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">新 密 码：</td>
    <td><input name="newpsw" type="password" id="newpsw" size="40" maxlength="30" class="input-text"> 6-30个字符</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">确认密码：</td>
    <td><input name="endpsw" type="password" id="endpsw" size="40" maxlength="30" class="input-text"> 6-30个字符</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">状态控制：</td>
    <td><input name="pass" type="radio" value="1" checked>开通　　<input name="pass" type="radio" value="0">关闭</td>
  </tr>
  <tr class="tdbg2">
    <td colspan="2"><input type="submit" class="bt" value="确定提交">　　<input type="reset" class="bt" value="重新填写"></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>