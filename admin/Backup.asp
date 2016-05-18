<!--#include file="chk.asp"-->
<%
call CloseConn()
sqlserver = "(local)"
sqlname = "diwei8_com"
sqlpassword = "<%65063874"
sqlLoginTimeout = 15
databasename = "diwei8_com"
bak_file = "D:\databack\"
%>
<table cellpadding="0" cellspacing="1" class="border">
  <tr><th colspan="2">备份数据库</th></tr>
  <form method="post">
  <input name="Oper" value="backup" type="hidden" />
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;当前数据库名称：<span class="red-b">diwei8_com.sql</span>（为了数据的安全，请定期备份好您的数据库，当天多次备份会被覆盖。）</td>
    <td><input type=submit value="备份数据库" class="bt" onClick="this.value='正在备份,请稍候...';this.disabled=true;form.submit();"></td>
  </tr>
  </form>
</table>
<br>
<table cellpadding="0" cellspacing="1" class="border">
  <tr>
    <th>已备份的数据列表</th>
    <th>备份时间</th>
    <th>删除</th>
    <th>还原</th>
  </tr>
<%
set MyFso=Server.CreateObject("Scripting.FileSystemObject")
Dim DataFolder,DataFileList,DataFile,DataFileName
If CheckDir(bak_file,1) = True Then 
else 
MakeNewsDir bak_file 
end if 
Set DataFolder=MyFso.GetFolder(bak_file)
Set DataFileList=DataFolder.Files
For Each DataFile in DataFileList
	'Temp=DataFile.DateCreated
	If Instr(DataFile,replace(databasename,".sql","")) Then
		DataFileName=DataFile.Name
%>
  <tr class="tdbg">
    <td class="t-center h25">&nbsp;<font color=red><%=DataFileName%></font></td>
    <td class="t-center"><%=DataFile.DateLastAccessed%></td>
    <form method="post" onSubmit="return confirm('确认要删除这个备份吗？')" />
    <input name="Oper" value="delfile" type="hidden" />
    <input name="FileName" value="<%=DataFileName%>" type="hidden" />
    <td class="t-center"><input type="submit" class="bt" value="删除" /></td>
    </form>
    <form method="post" onSubmit="return confirm('警告！！确认要还原数据吗？\n\n此操作将使网站内容回到备份之日的状态！备份日之后的数据将全部丢弃！\n\n谨慎操作！不建议进行此操作！')" />
    <input name="Oper" value="restore" type="hidden" />
    <input name="FileName" value="<%=DataFileName%>" type="hidden" />
    <td class="t-center"><input type="submit" class="bt" value="还原" /></td>
    </form>
  </tr>
<%
	End If
Next
SET MyFso=Nothing
%>
</table>
<%
DbOper = Request.form("Oper")
if(DbOper<>"")then
	'============备份数据库====================
	if(DbOper="backup")then
		Set srv=Server.CreateObject("SQLDMO.SQLServer")
		srv.LoginTimeout = sqlLoginTimeout
		srv.Connect sqlserver,sqlname, sqlpassword
		Set bak = Server.CreateObject("SQLDMO.Backup")
		bak.Database=databasename
		bak.Devices=Files
		bak.Files=bak_file&databasename&date()&".sql"
		On Error Resume Next
		bak.SQLBackup srv
		If err then
			response.write("<SCRIPT LANGUAGE=JavaScript>alert ('操作失败！网站使用繁忙中，请稍后再试！');window.location.href='Backup.asp';</script>")
			err.clear
			response.End
		End If
		response.write "<SCRIPT LANGUAGE=JavaScript>alert ('备份成功！');window.location.href='Backup.asp';</script>"
		response.end
	'============恢复数据库====================
	elseif(DbOper="restore")then
		Set srv=Server.CreateObject("SQLDMO.SQLServer")
		srv.LoginTimeout = sqlLoginTimeout
		srv.Connect sqlserver,sqlname, sqlpassword
		Set rest=Server.CreateObject("SQLDMO.Restore")
		rest.Action=0
		rest.Database=databasename
		rest.Devices=Files
		rest.Files=bak_file&Request.form("FileName")
		rest.ReplaceDatabase=True
		On Error Resume Next
		rest.SQLRestore srv
		If err then
			response.write("<SCRIPT LANGUAGE=JavaScript>alert ('操作失败！网站使用繁忙中，请稍后再试！');window.location.href='Backup.asp';</script>")
			err.clear
			response.End
		End If
		response.write "<SCRIPT LANGUAGE=JavaScript>alert ('恢复成功！');window.location.href='Backup.asp';</script>"
		response.end
	'============删除备份数据库====================
	elseif(DbOper="delfile")then
		if(CheckFileExists(bak_file&Request.form("FileName")))then
			DeleteFiles bak_file&Request.form("FileName")
			response.Redirect("Backup.asp")
			response.end
		else
			response.write "<SCRIPT LANGUAGE=JavaScript>alert ('备份文件【" & Request.form("FileName") & "】不存在！');window.location.href='Backup.asp';</script>"
			response.end
		end if
	end if
	set DbOper = Nothing
end if
%>
</body>
</html>