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
  <tr><th colspan="2">�������ݿ�</th></tr>
  <form method="post">
  <input name="Oper" value="backup" type="hidden" />
  <tr class="tdbg">
    <td class="t-left h25">&nbsp;��ǰ���ݿ����ƣ�<span class="red-b">diwei8_com.sql</span>��Ϊ�����ݵİ�ȫ���붨�ڱ��ݺ��������ݿ⣬�����α��ݻᱻ���ǡ���</td>
    <td><input type=submit value="�������ݿ�" class="bt" onClick="this.value='���ڱ���,���Ժ�...';this.disabled=true;form.submit();"></td>
  </tr>
  </form>
</table>
<br>
<table cellpadding="0" cellspacing="1" class="border">
  <tr>
    <th>�ѱ��ݵ������б�</th>
    <th>����ʱ��</th>
    <th>ɾ��</th>
    <th>��ԭ</th>
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
    <form method="post" onSubmit="return confirm('ȷ��Ҫɾ�����������')" />
    <input name="Oper" value="delfile" type="hidden" />
    <input name="FileName" value="<%=DataFileName%>" type="hidden" />
    <td class="t-center"><input type="submit" class="bt" value="ɾ��" /></td>
    </form>
    <form method="post" onSubmit="return confirm('���棡��ȷ��Ҫ��ԭ������\n\n�˲�����ʹ��վ���ݻص�����֮�յ�״̬��������֮������ݽ�ȫ��������\n\n������������������д˲�����')" />
    <input name="Oper" value="restore" type="hidden" />
    <input name="FileName" value="<%=DataFileName%>" type="hidden" />
    <td class="t-center"><input type="submit" class="bt" value="��ԭ" /></td>
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
	'============�������ݿ�====================
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
			response.write("<SCRIPT LANGUAGE=JavaScript>alert ('����ʧ�ܣ���վʹ�÷�æ�У����Ժ����ԣ�');window.location.href='Backup.asp';</script>")
			err.clear
			response.End
		End If
		response.write "<SCRIPT LANGUAGE=JavaScript>alert ('���ݳɹ���');window.location.href='Backup.asp';</script>"
		response.end
	'============�ָ����ݿ�====================
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
			response.write("<SCRIPT LANGUAGE=JavaScript>alert ('����ʧ�ܣ���վʹ�÷�æ�У����Ժ����ԣ�');window.location.href='Backup.asp';</script>")
			err.clear
			response.End
		End If
		response.write "<SCRIPT LANGUAGE=JavaScript>alert ('�ָ��ɹ���');window.location.href='Backup.asp';</script>"
		response.end
	'============ɾ���������ݿ�====================
	elseif(DbOper="delfile")then
		if(CheckFileExists(bak_file&Request.form("FileName")))then
			DeleteFiles bak_file&Request.form("FileName")
			response.Redirect("Backup.asp")
			response.end
		else
			response.write "<SCRIPT LANGUAGE=JavaScript>alert ('�����ļ���" & Request.form("FileName") & "�������ڣ�');window.location.href='Backup.asp';</script>"
			response.end
		end if
	end if
	set DbOper = Nothing
end if
%>
</body>
</html>