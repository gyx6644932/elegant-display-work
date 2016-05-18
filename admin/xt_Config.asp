<!--#include file="chk.asp"-->
<%
if request.Form("Oper") = "edit" then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "select * from xt_Config"
	rs.open sql,conn,1,3
	rs("SiteName")=ReplaceBadChar(request.Form("SiteName"))
	rs("d1")=ReplaceBadChar(request.Form("d1"))
	rs("d2")=ReplaceBadChar(request.Form("d2"))
	rs("flashNum")=request.Form("flashNum")
	rs("flashTime")=request.Form("flashTime")
	rs("siteltd")=ReplaceBadChar(request.Form("siteltd"))
	rs("addr")=ReplaceBadChar(request.Form("addr"))
	rs("sitetel")=ReplaceBadChar(request.Form("sitetel"))
	rs("sitefax")=ReplaceBadChar(request.Form("sitefax"))
	rs("youbian")=ReplaceBadChar(request.Form("youbian"))
	rs("beian")=ReplaceBadChar(request.Form("beian"))
	rs("mailform")=ReplaceBadChar(request.Form("mailform"))
	rs("mailusername")=ReplaceBadChar(request.Form("mailusername"))
	rs("mailuserpass")=request.Form("mailuserpass")
	rs("maildom")=ReplaceBadChar(request.Form("maildom"))
	rs("mailsmtp")=ReplaceBadChar(request.Form("mailsmtp"))
	rs("mailpop3")=ReplaceBadChar(request.Form("mailpop3"))
	rs("web_list")=request.Form("web_list")
	rs("web_fuwu")=request.Form("web_fuwu")
	rs.update
	rs.close
	call jsoutgo("配置成功！","xt_Config.asp")
end if
%>
<table cellpadding="0" cellspacing="1" class="border">
  <form method="post">
  <input type="hidden" name="Oper" value="edit">
  <tr><th colspan="2">网站基本信息</td></tr>
  <tr class="tdbg">
    <td class="t-right h25 w15">网站名称：</td>
    <td><input type="text" name="SiteName" value="<%=webc("SiteName")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">关 键 字：</td>
    <td><input type="text" name="d1" value="<%=webc("d1")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">网站描述：</td>
    <td><input type="text" name="d2" value="<%=webc("d2")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">公司名称：</td>
    <td><input type="text" name="siteltd" value="<%=webc("siteltd")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">公司地址：</td>
    <td><input type="text" name="addr" value="<%=webc("addr")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">联系电话：</td>
    <td><input type="text" name="sitetel" value="<%=webc("sitetel")%>" maxlength="45" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">传　　真：</td>
    <td><input type="text" name="sitefax" value="<%=webc("sitefax")%>" maxlength="45" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">邮政编码：</td>
    <td><input type="text" name="youbian" value="<%=webc("youbian")%>" maxlength="45" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">备案编号：</td>
    <td><input type="text" name="beian" value="<%=webc("beian")%>" maxlength="45" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">发件邮箱：</td>
    <td><input type="text" name="mailform" value="<%=webc("mailform")%>" maxlength="240" class="w100" title="提供网站系统给用户发送邮件，如用户进行找回密码操作，系统用此邮箱发送密码给用户，也可用来接收用户发给您的邮件，格式：*********@qq.com。"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">邮箱验证：</td>
    <td><input type="text" name="mailusername" value="<%=webc("mailusername")%>" maxlength="240" class="w100" title="系统调用邮箱时验证，126或163邮箱填邮箱用户名，即@之前部分！QQ及一些普通邮箱要写整个邮箱地址！"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">邮箱密码：</td>
    <td><input type="password" name="mailuserpass" value="<%=webc("mailuserpass")%>" maxlength="240" class="w100" title="提供正确的邮箱密码，网站系统才能使用发件系统！"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">邮箱域名：</td>
    <td><input type="text" name="maildom" value="<%=webc("maildom")%>" maxlength="240" class="w100" title="例如使用QQ邮箱填：mail.qq.com"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">邮箱SMTP：</td>
    <td><input type="text" name="mailsmtp" value="<%=webc("mailsmtp")%>" maxlength="30" class="w100" title="例如使用QQ邮箱必须在QQ邮箱首页帐户设置中开启SMTP！"></td>
  </tr>  
  <tr class="tdbg">
    <td class="t-right h25">邮箱POP3：</td>
    <td><input type="text" name="mailpop3" value="<%=webc("mailpop3")%>" maxlength="30" class="w100" title="例如使用QQ邮箱填：pop.qq.com"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">幻灯参数：</td>
    <td><select name="flashNum">
        <option value="0" <%if webc("flashNum")=0 then response.write "selected"%>>切换：透明过渡</option>
        <option value="1" <%if webc("flashNum")=1 then response.write "selected"%>>切换：运动透明</option>
        <option value="2" <%if webc("flashNum")=2 then response.write "selected"%>>切换：模糊过渡</option>
        <option value="3" <%if webc("flashNum")=3 then response.write "selected"%>>切换：运动模糊</option>
        </select><select name="flashTime">
        <option value="2" <%if webc("flashTime")=2 then response.write "selected"%>>间隔：2秒</option>
        <option value="3" <%if webc("flashTime")=3 then response.write "selected"%>>间隔：3秒</option>
        <option value="4" <%if webc("flashTime")=4 then response.write "selected"%>>间隔：4秒</option>
        <option value="5" <%if webc("flashTime")=5 then response.write "selected"%>>间隔：5秒</option>
        <option value="6" <%if webc("flashTime")=6 then response.write "selected"%>>间隔：6秒</option>
        <option value="7" <%if webc("flashTime")=7 then response.write "selected"%>>间隔：7秒</option>
        <option value="8" <%if webc("flashTime")=8 then response.write "selected"%>>间隔：8秒</option>
        <option value="9" <%if webc("flashTime")=9 then response.write "selected"%>>间隔：9秒</option>
        </select></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">首页套餐更多链接：</td>
    <td><input type="text" name="web_list" value="<%=webc("web_list")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">特惠服务更多链接：</td>
    <td><input type="text" name="web_fuwu" value="<%=webc("web_fuwu")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg2">
    <td colspan="2"><input type="submit" class="bt" value="确定更新">　　<input type="reset" class="bt" value="重新填写"></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>