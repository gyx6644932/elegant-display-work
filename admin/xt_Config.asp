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
	call jsoutgo("���óɹ���","xt_Config.asp")
end if
%>
<table cellpadding="0" cellspacing="1" class="border">
  <form method="post">
  <input type="hidden" name="Oper" value="edit">
  <tr><th colspan="2">��վ������Ϣ</td></tr>
  <tr class="tdbg">
    <td class="t-right h25 w15">��վ���ƣ�</td>
    <td><input type="text" name="SiteName" value="<%=webc("SiteName")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�� �� �֣�</td>
    <td><input type="text" name="d1" value="<%=webc("d1")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">��վ������</td>
    <td><input type="text" name="d2" value="<%=webc("d2")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">��˾���ƣ�</td>
    <td><input type="text" name="siteltd" value="<%=webc("siteltd")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">��˾��ַ��</td>
    <td><input type="text" name="addr" value="<%=webc("addr")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">��ϵ�绰��</td>
    <td><input type="text" name="sitetel" value="<%=webc("sitetel")%>" maxlength="45" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�������棺</td>
    <td><input type="text" name="sitefax" value="<%=webc("sitefax")%>" maxlength="45" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�������룺</td>
    <td><input type="text" name="youbian" value="<%=webc("youbian")%>" maxlength="45" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">������ţ�</td>
    <td><input type="text" name="beian" value="<%=webc("beian")%>" maxlength="45" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�������䣺</td>
    <td><input type="text" name="mailform" value="<%=webc("mailform")%>" maxlength="240" class="w100" title="�ṩ��վϵͳ���û������ʼ������û������һ����������ϵͳ�ô����䷢��������û���Ҳ�����������û����������ʼ�����ʽ��*********@qq.com��"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">������֤��</td>
    <td><input type="text" name="mailusername" value="<%=webc("mailusername")%>" maxlength="240" class="w100" title="ϵͳ��������ʱ��֤��126��163�����������û�������@֮ǰ���֣�QQ��һЩ��ͨ����Ҫд���������ַ��"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�������룺</td>
    <td><input type="password" name="mailuserpass" value="<%=webc("mailuserpass")%>" maxlength="240" class="w100" title="�ṩ��ȷ���������룬��վϵͳ����ʹ�÷���ϵͳ��"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">����������</td>
    <td><input type="text" name="maildom" value="<%=webc("maildom")%>" maxlength="240" class="w100" title="����ʹ��QQ�����mail.qq.com"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">����SMTP��</td>
    <td><input type="text" name="mailsmtp" value="<%=webc("mailsmtp")%>" maxlength="30" class="w100" title="����ʹ��QQ���������QQ������ҳ�ʻ������п���SMTP��"></td>
  </tr>  
  <tr class="tdbg">
    <td class="t-right h25">����POP3��</td>
    <td><input type="text" name="mailpop3" value="<%=webc("mailpop3")%>" maxlength="30" class="w100" title="����ʹ��QQ�����pop.qq.com"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�õƲ�����</td>
    <td><select name="flashNum">
        <option value="0" <%if webc("flashNum")=0 then response.write "selected"%>>�л���͸������</option>
        <option value="1" <%if webc("flashNum")=1 then response.write "selected"%>>�л����˶�͸��</option>
        <option value="2" <%if webc("flashNum")=2 then response.write "selected"%>>�л���ģ������</option>
        <option value="3" <%if webc("flashNum")=3 then response.write "selected"%>>�л����˶�ģ��</option>
        </select><select name="flashTime">
        <option value="2" <%if webc("flashTime")=2 then response.write "selected"%>>�����2��</option>
        <option value="3" <%if webc("flashTime")=3 then response.write "selected"%>>�����3��</option>
        <option value="4" <%if webc("flashTime")=4 then response.write "selected"%>>�����4��</option>
        <option value="5" <%if webc("flashTime")=5 then response.write "selected"%>>�����5��</option>
        <option value="6" <%if webc("flashTime")=6 then response.write "selected"%>>�����6��</option>
        <option value="7" <%if webc("flashTime")=7 then response.write "selected"%>>�����7��</option>
        <option value="8" <%if webc("flashTime")=8 then response.write "selected"%>>�����8��</option>
        <option value="9" <%if webc("flashTime")=9 then response.write "selected"%>>�����9��</option>
        </select></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">��ҳ�ײ͸������ӣ�</td>
    <td><input type="text" name="web_list" value="<%=webc("web_list")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�ػݷ���������ӣ�</td>
    <td><input type="text" name="web_fuwu" value="<%=webc("web_fuwu")%>" maxlength="240" class="w100"></td>
  </tr>
  <tr class="tdbg2">
    <td colspan="2"><input type="submit" class="bt" value="ȷ������">����<input type="reset" class="bt" value="������д"></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>