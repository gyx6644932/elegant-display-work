<%
Dim connstr,Conn,sql,rs
ConnStr="Provider=SQLOLEDB; User ID=jiuchang; Password=<%6644932; Initial CataLog=jiuchang; Data Source=(local);"
'On Error Resume Next
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.open ConnStr
If Err Then
	err.Clear
	Set Conn = Nothing
	Response.Write("<br><p align=center><font color='red'>�Բ��𣡣���վ���ڽ�������ά���������������Ӻ���ʹ�ñ�վ��лл������</font></p><br /><br />"&Err.Source&" ("&Err.Number&")")
	Response.End
End If
sub CloseConn()
	Conn.close
	set Conn=nothing
end sub
%>