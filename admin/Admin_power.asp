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
	Response.Write "<script>alert('���óɹ����µ�Ȩ�޷�����Ҫ�ѵ�¼���û����µ�¼����Ч��');window.location.href='Admin_power.asp?id="&id&"';</script>"
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
  <tr><th colspan="6">Ȩ������(�û���<%=rs("Userid")%>)</th></tr>
  <tr class="tdbg">
    <td colspan="6" class="t-left h25"><input type="checkbox" <%if adminpower(1) then response.write "checked"%> name="power1" value="1"><strong>ϵͳ����</strong></td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(2) then response.write "checked"%> name="power2" value="1">��վ����Ա</td>
    <td><input type="checkbox" <%if adminpower(3) then response.write "checked"%> name="power3" value="1">���ݿ����</td>
    <td><input type="checkbox" <%if adminpower(4) then response.write "checked"%> name="power4" value="1">�ϴ��ļ�����</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
  </tr>
  <tr class="tdbg"><td colspan="6" class="t-left h25">&nbsp;</td></tr>
  <tr class="tdbg">
    <td colspan="6" class="t-left h25"><input type="checkbox" <%if adminpower(5) then response.write "checked"%> name="power5" value="1"><strong>��վ����</strong></td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(6) then response.write "checked"%> name="power6" value="1">��������</td>
    <td><input type="checkbox" <%if adminpower(7) then response.write "checked"%> name="power7" value="1">��ҳ����</td>
    <td><input type="checkbox" <%if adminpower(8) then response.write "checked"%> name="power8" value="1">�õƹ���</td>
    <td><input type="checkbox" <%if adminpower(9) then response.write "checked"%> name="power9" value="1">��������</td>
	<td><input type="checkbox" <%if adminpower(10) then response.write "checked"%> name="power10" value="1">�����˵�</td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(11) then response.write "checked"%> name="power11" value="1">����ͨ��</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
  </tr>
  <tr class="tdbg"><td colspan="6" class="t-left h25">&nbsp;</td></tr>
  <tr class="tdbg">
    <td colspan="6" class="t-left h25"><input type="checkbox" <%if adminpower(12) then response.write "checked"%> name="power12" value="1"><strong>��վ�����ײ͹���</strong></td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(13) then response.write "checked"%> name="power13" value="1">�ײ͹���</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
  </tr>
  <tr class="tdbg"><td colspan="6" class="t-left h25">&nbsp;</td></tr>
  <tr class="tdbg">
    <td colspan="6" class="t-left h25"><input type="checkbox" <%if adminpower(14) then response.write "checked"%> name="power14" value="1"><strong>�ػݷ������</strong></td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(15) then response.write "checked"%> name="power15" value="1">�ػݷ������</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr class="tdbg"><td colspan="6" class="t-left h25">&nbsp;</td></tr>
  <tr class="tdbg">
    <td colspan="6" class="t-left h25"><input type="checkbox" <%if adminpower(16) then response.write "checked"%> name="power16" value="1"><strong>���Ź���</strong></td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(17) then response.write "checked"%> name="power17" value="1">�������</td>
    <td><input type="checkbox" <%if adminpower(18) then response.write "checked"%> name="power18" value="1">���Ź���</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr class="tdbg"><td colspan="6" class="t-left h25">&nbsp;</td></tr>
  <tr class="tdbg">
    <td colspan="6" class="t-left h25"><input type="checkbox" <%if adminpower(19) then response.write "checked"%> name="power19" value="1"><strong>����ͳ��</strong></td>
  </tr>
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;</td>
    <td><input type="checkbox" <%if adminpower(20) then response.write "checked"%> name="power20" value="1">��������</td>
    <td><input type="checkbox" <%if adminpower(21) then response.write "checked"%> name="power21" value="1">��վ�ƹ�</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
  </tr>
  <tr class="tdbg"><td colspan="6" class="t-left h25">&nbsp;</td></tr>
  <tr class="tdbg2">
    <td colspan="6"><input name="Submit" type="submit" class="bt" value="ȷ���ύ">����<input name="Submit" type="reset" class="bt" value="������д"></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>