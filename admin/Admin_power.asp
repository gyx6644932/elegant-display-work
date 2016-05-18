<!--#include file="chk.asp"-->
<%
id = Request.QueryString("id")
if Request.Form("Oper") = "edit" then
	set rs=server.createobject("adodb.recordset")
	sql="select * from admin where id = "&id
	rs.open sql,conn,1,3
	s="0|"
		for i=1 to 99
			if Request("power"&i) <> "" then
				s=s & "1|"
			else
				s=s & "0|"
			end if
		next
	rs("adminpower")=s & "0"
	rs.update
	rs.close
	set rs=nothing
	Response.Write "<script>alert('配置成功！新的权限方案需要已登录的用户重新登录才有效！');window.location.href='Admin_power.asp?id="&id&"';</script>"
	Response.End()
end if
set rs=server.createobject("adodb.recordset")
rs.open "select * from admin where id="& id,conn,1,1
if rs.eof then
	response.end
else
	dim adminpower(100)
	s=split(rs("adminpower"),"|")
	For i=0 to UBound(s)
		adminpower(i)=CBool(s(i))
	Next
end if
%>
<!--#include file="Admin_top.asp"-->
<table cellpadding="0" cellspacing="1" class="border">
  <form method="post">
  <input type="hidden" name="id" value="<%=rs("id")%>">
  <input type="hidden" name="Oper" value="edit">
  <tr><th colspan="6">权限设置(用户：<%=rs("Userid")%>)</th></tr>
  <tr class="tdbg">
    <td colspan="6" class="t-left h25"><input type="checkbox" <%if adminpower(1) then response.write "checked"%> name="power1" value="1"><strong>系统管理</strong></td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(2) then response.write "checked"%> name="power2" value="1">网站管理员</td>
    <td><input type="checkbox" <%if adminpower(3) then response.write "checked"%> name="power3" value="1">数据库管理</td>
    <td><input type="checkbox" <%if adminpower(4) then response.write "checked"%> name="power4" value="1">上传文件管理</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
  </tr>
  <tr class="tdbg"><td colspan="6" class="t-left h25">&nbsp;</td></tr>
  <tr class="tdbg">
    <td colspan="6" class="t-left h25"><input type="checkbox" <%if adminpower(5) then response.write "checked"%> name="power5" value="1"><strong>网站管理</strong></td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(6) then response.write "checked"%> name="power6" value="1">基本配置</td>
    <td><input type="checkbox" <%if adminpower(7) then response.write "checked"%> name="power7" value="1">单页配置</td>
    <td><input type="checkbox" <%if adminpower(8) then response.write "checked"%> name="power8" value="1">幻灯管理</td>
    <td><input type="checkbox" <%if adminpower(9) then response.write "checked"%> name="power9" value="1">友情链接</td>
	<td><input type="checkbox" <%if adminpower(10) then response.write "checked"%> name="power10" value="1">导航菜单</td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(11) then response.write "checked"%> name="power11" value="1">快速通道</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
  </tr>
  <tr class="tdbg"><td colspan="6" class="t-left h25">&nbsp;</td></tr>
  <tr class="tdbg">
    <td colspan="6" class="t-left h25"><input type="checkbox" <%if adminpower(12) then response.write "checked"%> name="power12" value="1"><strong>网站建设套餐管理</strong></td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(13) then response.write "checked"%> name="power13" value="1">套餐管理</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
  </tr>
  <tr class="tdbg"><td colspan="6" class="t-left h25">&nbsp;</td></tr>
  <tr class="tdbg">
    <td colspan="6" class="t-left h25"><input type="checkbox" <%if adminpower(14) then response.write "checked"%> name="power14" value="1"><strong>特惠服务管理</strong></td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(15) then response.write "checked"%> name="power15" value="1">特惠服务管理</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr class="tdbg"><td colspan="6" class="t-left h25">&nbsp;</td></tr>
  <tr class="tdbg">
    <td colspan="6" class="t-left h25"><input type="checkbox" <%if adminpower(16) then response.write "checked"%> name="power16" value="1"><strong>新闻管理</strong></td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(17) then response.write "checked"%> name="power17" value="1">添加新闻</td>
    <td><input type="checkbox" <%if adminpower(18) then response.write "checked"%> name="power18" value="1">新闻管理</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr class="tdbg"><td colspan="6" class="t-left h25">&nbsp;</td></tr>
  <tr class="tdbg">
    <td colspan="6" class="t-left h25"><input type="checkbox" <%if adminpower(19) then response.write "checked"%> name="power19" value="1"><strong>访问统计</strong></td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(20) then response.write "checked"%> name="power20" value="1">流量数据</td>
    <td><input type="checkbox" <%if adminpower(21) then response.write "checked"%> name="power21" value="1">网站推广</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
  </tr>
  <tr class="tdbg"><td colspan="6" class="t-left h25">&nbsp;</td></tr>
  <tr class="tdbg2">
    <td colspan="6"><input name="Submit" type="submit" class="bt" value="确定提交">　　<input name="Submit" type="reset" class="bt" value="重新填写"></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>