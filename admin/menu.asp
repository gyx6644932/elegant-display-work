<!--#include file="chk.asp"-->
<table cellpadding="0" cellspacing="1" class="border">
  <tr><th>导航菜单管理</th></tr>
  <tr class="tdbg"><td class="t-left h25"><a href="menu_edit.asp?FB_iframe=true&height=100&width=300" title="新增导航" class="firstebox">新增导航</a></td></tr>
</table>
<br>
<table cellpadding="0" cellspacing="1" class="border">
  <tr>
    <th>排序</th>
    <th>链接名称</th>
    <th>链接地址</th>
    <th>修改</th>
    <th>删除</th>
  </tr>
<%
set rs = server.createobject("adodb.recordset")
sql = "select * from xt_menu order by px asc,id asc"
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
    <td class="t-center h25"><a href="Desc.asp?id=<%=rs("id")%>&mssql=xt_menu&FB_iframe=true&height=50&width=210" class="firstebox" title="修改排序值"><%=rs("px")%></a></td>
    <td class="t-center"><%=rs("title")%></td>
    <td class="t-center"><%=rs("content")%></td>
    <td class="t-center"><a href="menu_edit.asp?id=<%=rs("id")%>&FB_iframe=true&height=100&width=300" title="修改菜单" class="firstebox"><img src="images/edit/edit.gif"></a></td>
    <td class="t-center"><a href="Del.asp?id=<%=rs("id")%>&mssql=xt_menu&FB_iframe=true&height=100&width=200" class="firstebox" title="删除确认"><img src="images/edit/adel.gif"></a></td>
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