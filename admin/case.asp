<!--#include file="Chk.asp"-->
<%
bid = go_num(Request.QueryString("bid"),0)
typejj = go_num(Request.QueryString("typejj"),1)
Keyword = ReplaceBadChar(Request.QueryString("Keyword"))
%>
<table cellpadding="0" cellspacing="1"  class="border">
  <tr><th>��������</th></tr>
  <form method="get">
  <tr class="tdbg">
    <td class="t-left h25"><input type="button" value="��������" title="��������" class="firstebox bt" alt="case_edit.asp?FB_iframe=true&height=0.8&width=740" /> <select name="bid">
<option value="0">ȫ������</option>
<%
set rs = server.CreateObject("adodb.recordset")
sql = "Select * from case_bid order by px asc,id desc"
rs.Open sql,Conn,1,1
if not rs.eof then
do while not rs.eof
%>
<option value="<%=rs("id")%>" <%if Cstr(rs("id")) = Cstr(bid) then Response.Write("selected='selected'")%>><%=rs("title")%></option>
<%
rs.movenext
loop
end if
rs.close
set rs=nothing
%></select> <select name="typejj">
<option value="1" <%if typejj = 1 then Response.Write("selected='selected'")%>>������������վ����</option>
<option value="2" <%if typejj = 2 then Response.Write("selected='selected'")%>>������������վ����</option>
</select> <input type="text" name="Keyword" value="<%=Keyword%>" style="width:150px;" /> <input type="submit" class="bt" value="����" /></td>
  </tr>
  </form>
</table><br />
<table cellpadding="0" cellspacing="1" class="border">
  <tr>
    <th>����</th>
    <th>����</th>
    <th>��վ����</th>
    <th>����ͼ</th>
    <th>����ʱ��</th>
    <th>�޸�</th>
	<th>ɾ��</th>
  </tr> 
<%
set rs = Server.CreateObject("ADODB.Recordset")
sql = "Select * From [case] Where 1 = 1"
if bid <> 0 then sql = sql&" and bid = "&bid
if Keyword <> "" then
	if typejj = 1 then
		sql = sql&" and title like '%"&Keyword&"%'"
	elseif typejj = 2 then
		sql = sql&" and content like '%"&Keyword&"%'"
	end if
end if
sql = sql&" Order By px asc,id Desc"
rs.Open sql,Conn,1,1
if not rs.EOF then
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
  <tr class="tdbg">
    <td class="t-center h25"><%=rs("px")%></td>
    <td class="t-center"><%=rs("bname")%></td>
    <td class="t-center"><a href="<%=rs("url")%>" target="_blank"><%=rs("title")%></a></td>
    <td class="t-center"><a href="<%=rs("img")%>" title="<%=rs("title")%>" rel="pic" class="firstebox"><img src="<%=rs("img")%>" style="width:50px;height:30px;" /></a></td>
    <td class="t-center"><%=rs("addtime")%></td>
    <td class="t-center"><a href="case_edit.asp?id=<%=rs("id")%>&FB_iframe=true&height=0.8&width=740" class="firstebox" title="�����޸�"><img src="/images/edit/edit.gif"></a></td>
	<td class="t-center"><a href="Del.asp?id=<%=rs("id")%>&mssql=case&FB_iframe=true&height=100&width=200" class="firstebox" title="ɾ��ȷ��"><img src="/images/edit/adel.gif"></a></td>
  </tr>
<%
	rs.movenext
	Next
End If
rs.close
set rs = nothing
%>
    <tr class="tdbg2"><td colspan="7"><%Call PageControl(iCount, maxpage, page, iPageSize)%></td></tr>
</table>
<%call CloseConn()%>
</body>
</html>