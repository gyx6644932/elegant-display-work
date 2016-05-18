<%
set webc = server.CreateObject("ADODB.recordset")
webc.open "select * from xt_config",conn,1,1
%>