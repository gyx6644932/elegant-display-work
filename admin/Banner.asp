<!--#include file="chk.asp"-->
<table cellpadding="0" cellspacing="1" class="border">
  <tr><th>幻灯管理</th></tr>
  <tr class="tdbg"><td class="t-left h25"><a href="Banner_edit.asp?FB_iframe=true&height=510&width=510" title="添加幻灯" class="firstebox">添加幻灯</a></td></tr>
</table>
<br>
<table cellpadding="0" cellspacing="1" class="border">
  <tr>
    <th style="width:1%;">图片预览</th>
    <th>幻灯名称</th>
    <th>排序</th>
    <th>链接地址</th>
    <th>修改</th>
    <th>删除</th>
  </tr>
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "Select * From flash Order By px asc,id Desc"
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
	if instr(rs("imgpath"),"http://") > 0 then
		imgurl = rs("imgpath")
	else
		imgurl = "../"&rs("imgpath")
	end if
%>
  <tr class="tdbg" onMouseOver="this.className='tdbg3'" onMouseOut="this.className='tdbg'">
    <td style="padding:0px;"><img src="<%=imgurl%>" width="200" height="40" /></td>
    <td class="t-center"><%=rs("title")%></td>
    <td class="t-center"><a href="Desc.asp?id=<%=rs("id")%>&mssql=flash&FB_iframe=true&height=50&width=210" class="firstebox" title="修改排序值"><%=rs("px")%></a></td>
    <td class="t-left"><%=rs("link")%></td>
    <td class="t-center"><a href="banner_edit.asp?id=<%=rs("id")%>&FB_iframe=true&height=510&width=510" title="修改幻灯" class="firstebox"><img src="images/edit/edit.gif"></a></td>
    <td class="t-center"><a href="del.asp?id=<%=rs("id")%>&mssql=flash&FB_iframe=true&height=100&width=200" class="firstebox" title="删除确认"><img src="images/edit/adel.gif"></a></td>
  </tr>
<%
rs.movenext
Next
End If
%>
  <tr class="tdbg2">
    <td colspan="6"><%Call PageControl(iCount, maxpage, page, iPageSize)%></td>
  </tr>
</table>
<%call CloseConn()%>
</body>
</html>