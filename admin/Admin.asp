<!--#include file="chk.asp"-->
<%
id = Request.QueryString("id")
if request.QueryString("action")="passoff" then
	conn.execute("Update admin Set pass=0 where id = "&id)
	Response.Redirect("admin.asp")
	Response.End()
end if
if request.QueryString("action")="passon" then
	conn.execute("Update admin Set pass=1 where id = "&id)
	Response.Redirect("Admin.asp")
	Response.End()
end if
if request.Form("Oper") = "delall" then
	id = request.Form("id")
	if id <> "" then
		id=split(id,",")
		for i=0 to UBound(id)
		conn.execute("delete from admin where id = "&id(i))
		next
	end if
	Response.Redirect("Admin.asp")
	Response.End()
end if
%>
<script>
function CheckAll(form){
	for (var i=0;i<form.elements.length;i++){
		var e = form.elements[i];
		if (e.name != 'chkall'){e.checked = form.chkall.checked;}
	}
}
</script>
<%
dim adminpower(100)
s=split(Request.Cookies("adminpower"),"|")
For i=0 to UBound(s)
	adminpower(i)=CBool(s(i))
Next
%>
<table cellpadding="0" cellspacing="1" class="border">
  <tr><th>����Ա����</th></tr>
  <tr class="tdbg"><td class="t-left h25"><a href="Admin.asp">�������Ա</a> | <%if adminpower(99) then%><a href="Admin_add.asp">��������Ա</a><%else%><span class="ccc">��������Ա</span><%end if%></td></tr>
</table>
<br>
<table cellpadding="0" cellspacing="1" class="border">
  <form name="delform" method="post">
  <input type="hidden" name="Oper" value="delall">
  <tr>
    <th><input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)"></th>
    <th>�ʺ�</th>
    <th>���</th>
    <th>��¼</th>
    <th>����¼ʱ��</th>
    <th>����¼IP</th>
    <th>״̬</th>
    <th>�޸�</th>
    <th>Ȩ��</th>
  </tr>
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from admin where userkey = 0 order by userkey desc,id asc"
rs.Open sql, conn, 1, 1
If not (rs.bof and rs.EOF) Then
    rs.PageSize = 20 
    iCount = rs.RecordCount 
    iPageSize = rs.PageSize
    maxpage = rs.PageCount
    page = request("page")
    If Not IsNumeric(page) Or page = "" Then
        page = 1
    Else
        page = CInt(page)
    End If
    If page<1 Then
        page = 1
    ElseIf page>maxpage Then
        page = maxpage
    End If
    rs.AbsolutePage = Page
    If page = maxpage Then
        x = iCount - (maxpage -1) * iPageSize
    Else
        x = iPageSize
    End If
    For i = 1 To x
%>
  <tr class="tdbg" onMouseOver="this.className='tdbg3'" onMouseOut="this.className='tdbg'">
    <td class="t-center h25"><input type="checkbox" name="ID" value="<%=rs("id")%>" <%if rs("userid")=Request.Cookies("admin")("user") then response.Write" disabled" end if%>></td>
    <td class="t-center"><%=rs("userid")%></td>
    <td class="t-center"><%=rs("title")%></td>
    <td class="t-center"><%=rs("gonum")%>&nbsp;��</td>
    <td class="t-center"><%if rs("lasttime")="���� 12:00:00" then%>--<%else%><%=rs("lasttime")%><%end if%></td>
    <td class="t-center"><%if rs("goip")="0" then%>--<%else%><%=rs("goip")%><%end if%></td>
    <td class="t-center"><%if adminpower(99) then%><%if rs("userid")=Request.Cookies("admin")("user") then%>--<%else%><%if rs("pass")=1 then%><a href="Open.asp?mssql=admin&mscode=pass&Oper=go_off&id=<%=rs("id")%>&FB_iframe=true" class="firstebox" title="����ر�"><img src="images/edit/unlock.gif"></a><%else%><a href="Open.asp?mssql=admin&mscode=pass&Oper=go_on&id=<%=rs("id")%>&FB_iframe=true" class="firstebox" title="�����ͨ"><img src="images/edit/lock.gif"></a><%end if%><%end if%><%else%><img src="images/edit/unopen.gif"><%end if%></td>
    <td class="t-center"><%if adminpower(99) and Cstr(rs("userid")) <> Cstr(Request.Cookies("admin")("user")) then%><a href="Admin_pass.asp?id=<%=rs("id")%>"><img src="images/edit/edit.gif"></a><%else%><img src="images/edit/unedit.gif"><%end if%></td>
    <td class="t-center"><%if adminpower(99) then%><a href="Admin_power.asp?id=<%=rs("id")%>"><img src="images/edit/come.gif"></a><%else%><img src="images/edit/uncome.gif"><%end if%></td>
  </tr>
<%
rs.movenext
Next
End If
%>
  </form>
  <tr class="tdbg2">
    <td colspan="9">
      <table cellpadding="0" cellspacing="0" style="width:100%;border:none;">
        <tr>
          <td style="width:0%;"><%if adminpower(99) then%><input type="button" onClick="if(confirm('ȷ��Ҫɾ��ѡ�е���Ϣ��һ��ɾ�������ָܻ���'))delform.submit()" class="bt" value="ɾ��ѡ��"><%else%><input type="button" class="bt" value="ɾ��ѡ��" disabled><%end if%></td>
          <td><%Call PageControl(iCount, maxpage, page, iPageSize)%></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%call CloseConn()%>
</body>
</html>