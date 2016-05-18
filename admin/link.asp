<!--#include file="chk.asp"-->
<table cellpadding="0" cellspacing="1" class="border">
  <tr><th>友情链接管理</th></tr>
  <tr class="tdbg"><td class="t-left h25"><a href="link_edit.asp?FB_iframe=true&height=100&width=400" title="新增链接" class="firstebox">新增链接</a></td></tr>
</table>
<br>
<table cellpadding="0" cellspacing="1" class="border">
  <tr>
    <th>排序</th>
    <th>网站名称</th>
    <th>网址</th>
    <th>修改</th>
    <th>删除</th>
  </tr>
<%
set rs = server.createobject("adodb.recordset")
sql = "select * from link order by px asc,id asc"
rs.open sql,conn,1,1
if not rs.eof then
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
    <td class="t-center h25"><a href="Desc.asp?id=<%=rs("id")%>&mssql=link&FB_iframe=true&height=50&width=210" class="firstebox" title="修改排序值"><%=rs("px")%></a></td>
    <td class="t-center"><%=rs("title")%></td>
    <td class="t-center"><%=rs("url")%></td>
    <td class="t-center"><a href="link_edit.asp?id=<%=rs("id")%>&FB_iframe=true&height=100&width=400" title="修改链接" class="firstebox"><img src="images/edit/edit.gif"></a></td>
    <td class="t-center"><a href="Del.asp?id=<%=rs("id")%>&mssql=link&FB_iframe=true&height=100&width=200" class="firstebox" title="删除确认"><img src="images/edit/adel.gif"></a></td>
  </tr>
<%
	rs.movenext 
	Next
End If
%>
  <tr class="tdbg2"><td colspan="5"><%Call PageControl(iCount, maxpage, page, iPageSize)%></td></tr>
</table>
<%call CloseConn()%>
</body>
</html>