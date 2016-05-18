<!--#include file="Chk.asp"-->
<table cellpadding="0" cellspacing="1" class="border">
  <tr><th>线路一级分类管理</th></tr>
  <tr class="tdbg"><td class="h25"><input type="button" value="新增分类" title="新增分类" class="firstebox bt" alt="case_bid_edit.asp?FB_iframe=true&height=73&width=300" /></td></tr>
</table><br />
<table cellpadding="0" cellspacing="1" class="border">
  <tr>
    <th>排序</th>
    <th>分类名称</th>
    <th>修改</th>
    <th>删除</th>
  </tr>
<%
set rs = server.createobject("adodb.recordset")
sql = "select * from case_bid order by px asc,id asc"
rs.open sql,conn,1,1
set sql = nothing
dim iCount,maxpage,page,iPageSize,x
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
  <tr class="tdbg">
    <td class="t-center h25"><%=rs("px")%></td>
    <td class="t-center"><%=rs("title")%></td>
    <td class="t-center"><a href="case_bid_edit.asp?id=<%=rs("id")%>&FB_iframe=true&height=73&width=300" title="修改分类" class="firstebox"><img src="/images/edit/edit.gif"></a></td>
    <td class="t-center"><a href="Del.asp?id=<%=rs("id")%>&mssql=case_bid&FB_iframe=true&height=100&width=200" class="firstebox" title="删除确认"><img src="/images/edit/adel.gif"></a></td>
  </tr>
<%
	rs.movenext:Next
End If
rs.close
set rs = nothing
%>
  <tr class="tdbg2"><td colspan="4"><%Call PageControl(iCount, maxpage, page, iPageSize)%></td></tr>
</table>
<%call CloseConn()%>
</body>
</html>