<!--#include file="chk.asp"-->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
if Request.QueryString("folder") <> "" and not isnull(Request.QueryString("folder")) then
	UploadDir="../Uploadfile/flash/"&Request.QueryString("folder")&"/"
else
	UploadDir="../Uploadfile/flash/"
end if
TruePath=Server.MapPath(UploadDir)
If not IsObjInstalled("Scripting.FileSystemObject") Then
	Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
Else
	set fso=CreateObject("Scripting.FileSystemObject")
	if request("Action")="Del" then
		whichfile=server.mappath(Request("FileName"))
		if fso.folderExists(whichfile) then
			fso.deletefolder(whichfile)
		else
			Set thisfile = fso.GetFile(whichfile) 
			thisfile.Delete True
		end if
		Response.Write("<script>window.history.go(-1);</script>")
		Response.End()
	end if
	if fso.FolderExists(TruePath)then
		FileCount=0
		TotleSize=0
		Set theFolder=fso.GetFolder(TruePath)
		for each objfoldercount in theFolder.subfolders
			FileCount = FileCount + 1
			TotleSize = TotleSize + objfoldercount.Size
		next
		For Each theFile In theFolder.Files
			FileCount = FileCount + 1
			TotleSize = TotleSize + theFile.Size
		next
	else
		Response.Write("<script>alert('找不到目标存储文件夹！可能是配置有误！');window.location.href='Javascript:history.back()';</script>")
		Response.End()
	end if
end if
%>
<script language="JavaScript">
function ConfirmDel()
{
if (confirm("你真的要删除此文件吗!"))
	return true;
else
	return false;
}
</script>
<!--#include file="Upload_top.asp"-->
<table cellpadding="0" cellspacing="1" class="border">
  <tr>
    <th>文件名</th>
    <th>缩略图</th>
    <th>文件大小</th>
    <th>文件类型</th>
    <th>最后修改时间</th>
    <th>删除</th>
  </tr>
<%
iPageSize = 20
iCount = FileCount
if (iCount mod iPageSize)=0 then
	maxpage = iCount \ iPageSize
else
	maxpage = iCount \ iPageSize + 1
end if
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
If page = maxpage Then
    x = iCount - (maxpage -1) * iPageSize
Else
    x = iPageSize
End If
Thepagesize = 0
Themove = 0
for each objfoldercount in theFolder.subfolders
	Themove = Themove + 1
	if Themove > (iPageSize * Page) then
		exit for
	elseif Themove > iPageSize * (Page - 1) then
%>
  <tr class="tdbg" onMouseOver="this.className='tdbg3'" onMouseOut="this.className='tdbg'">
    <td class="t-center h25"><strong>&nbsp;<%=objfoldercount.Name%></strong></td>
    <td class="t-center"><a href="?folder=<%=objfoldercount.Name%>" title="<%=objfoldercount.Name%>"><img src="../html/plugins/filemanager/images/folder-64.gif" border="0" /></a></td>
    <td class="t-center"><%=objfoldercount.size%>字节&nbsp;</td>
    <td class="t-center"><%=objfoldercount.type%></td>
    <td class="t-center"><%=objfoldercount.DateCreated%></td>
    <td class="t-center"><a href="?Action=Del&FileName=<%=UploadDir&objfoldercount.Name%>" onClick="return ConfirmDel()"><img src="images/edit/adel.gif"></a></td>
  </tr>
<%
		Thepagesize = Thepagesize + objfoldercount.Size
	end if
next
For Each theFile In theFolder.Files
	Themove = Themove + 1
	if Themove > (iPageSize * Page) then
		exit for
	elseif Themove > iPageSize * (Page - 1) then
%>
  <tr class="tdbg" onMouseOver="this.className='tdbg3'" onMouseOut="this.className='tdbg'">
    <td class="t-center h25"><strong>&nbsp;<%=theFile.Name%></strong></td>
    <td class="t-center"><embed src="<%=UploadDir&theFile.Name%>" wmode="transparent" quality="high" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="80" height="80"></embed></td>
    <td class="t-center"><%=theFile.size%>字节&nbsp;</td>
    <td class="t-center"><%=theFile.type%></td>
    <td class="t-center"><%=theFile.DateLastModified%></td>
    <td class="t-center"><a href="?Action=Del&FileName=<%=UploadDir&theFile.Name%>" onClick="return ConfirmDel()"><img src="images/edit/adel.gif"></a></td>
  </tr>
<%
		Thepagesize = Thepagesize + theFile.Size
	end if
Next
%>
  <tr class="tdbg2"><td colspan="6">&nbsp;本页文件占用 <strong><%=round(Thepagesize / 1024 / 1024,2)%></strong> MB (<strong><%=Thepagesize%></strong> 字节) / 所有文件占用 <strong><%=round(TotleSize / 1024 / 1024,2)%></strong> MB (<strong><%=TotleSize%></strong> 字节)</td></tr>
  <tr class="tdbg2"><td colspan="6"><%Call PageControl(iCount, maxpage, page, iPageSize)%></td></tr>
</table>
<%call CloseConn()%>
</body>
</html>