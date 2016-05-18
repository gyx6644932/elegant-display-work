<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="../inc/config.asp"-->
<%
if replace(Request.Cookies("admin")("user"),"'","")=""  then
	Response.write "<script language='javascript'>window.top.location.href='login.asp';</script>"
	response.end
end if
set rsRm=server.createobject("adodb.recordset")
sql="select * from Admin where userid='"&Request.Cookies("admin")("user")&"'"
rsRm.open sql,conn,1,1
if rsRm.bof and rsRm.eof then
	response.write "<script LANGUAGE='javascript'>alert('本站拒绝远程登录！');window.top.location.href='logout.asp';</script>"
	response.End()
else
	if md5(md5(rsRm("sj_no"))&rsRm("userpsw"))<>Request.Cookies("admin")("psw") then
		Response.write "<script language='javascript'>window.top.location.href='logout.asp';</script>"
		response.end
	end if
end if
rsRm.close
set rsRm=nothing
dim ComeUrl,cUrl
ComeUrl=lcase(trim(request.ServerVariables("HTTP_REFERER")))
'===================================================================
'IE为6.0时开
'===================================================================
if ComeUrl<>"" then
	cUrl=trim("http://" & Request.ServerVariables("SERVER_NAME"))
	if mid(ComeUrl,len(cUrl)+1,1)=":" then
		cUrl=cUrl & ":" & Request.ServerVariables("SERVER_PORT")
	end if
	cUrl=lcase(cUrl & request.ServerVariables("SCRIPT_NAME"))
	if lcase(left(ComeUrl,instrrev(ComeUrl,"/")))<>lcase(left(cUrl,instrrev(cUrl,"/"))) then
		response.write "<br><p align=center><font color='red'>警告！！不允许从外部链接地址访问本系统的后台管理页面，您的IP等信息被记录！</font></p>"
		response.end
	end if
else
	response.write "<br><p align=center><font color='red'>警告！！不允许直接输入地址访问本系统的后台管理页面。</font></p>"
	response.end
end if
'======================================================================
%>
<!--#include file="../inc/Function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/Admin.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.min.js"></script>
<script type="text/javascript" src="/js/admin_tr.js"></script>
<script type="text/javascript" src="/js/jquery.firstebox.pack.js"></script>
<script type="text/javascript" src="/js/global.js"></script>
<style type="text/css">@import "../css/firstebox.css";</style>
<script type="text/javascript" src="/datepicker/WdatePicker.js"></script>
</HEAD>
<BODY>