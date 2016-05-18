<!--#include file="chk.asp"-->
<SCRIPT type="text/javascript"> 
<!--
function Onclicklink(url) { 
	window.clipboardData.setData("Text",url);
	alert("已复制");
} 
-->
</SCRIPT>
<script type="text/javascript">
<!--
function CheckAll(form) {
	for (var i=0;i<form.elements.length;i++) {
		var e = form.elements[i];
		if (e.name != 'chkall') e.checked = form.chkall.checked; 
	}
}
-->
</script>
<!--#include file="Web_top.asp"-->
<table cellpadding="0" cellspacing="1" class="border">
  <tr>
    <th>单页名称</th>
    <th>调用地址</th>
    <th>修改</th>
	<th>删除</th>
  </tr>
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from web order by id asc"
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
    <td class="t-center h25"><%=rs("title")%></td>
    <td class="t-center"><%if rs("bid") = 1 then%><A href="../index/Web_full.asp?id=<%=rs("id")%>" target="_blank">Web_full.asp?id=<%=rs("id")%></A>【<A href="javascript:Onclicklink('Web_full.asp?id=<%=rs("id")%>');">复制</A>】<%else%><A href="../index/Web.asp?id=<%=rs("id")%>" target="_blank">Web.asp?id=<%=rs("id")%></A>【<A href="javascript:Onclicklink('Web.asp?id=<%=rs("id")%>');">复制</A>】<%end if%></td>
    <td class="t-center"><a href="Web_edit.asp?id=<%=rs("id")%>"><img src="images/edit/edit.gif"></a></td>
	<td class="t-center"><a href="Del.asp?id=<%=rs("id")%>&mssql=web&FB_iframe=true&height=100&width=200" class="firstebox" title="删除确认"><img src="images/edit/adel.gif"></a></td>
  </tr>
<%
rs.movenext
Next
End If
%>
  <tr class="tdbg2">
    <td colspan="4"><%Call PageControl(iCount, maxpage, page, iPageSize)%></td>
  </tr>
</table>
<%call CloseConn()%>
</body>
</html>