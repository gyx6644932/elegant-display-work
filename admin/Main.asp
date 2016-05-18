<!--#include file="chk.asp"-->
<%
Dim theInstalledObjects(17)
theInstalledObjects(0) = "MSWC.AdRotator"
theInstalledObjects(1) = "MSWC.BrowserType"
theInstalledObjects(2) = "MSWC.NextLink"
theInstalledObjects(3) = "MSWC.Tools"
theInstalledObjects(4) = "MSWC.Status"
theInstalledObjects(5) = "MSWC.Counters"
theInstalledObjects(6) = "IISSample.ContentRotator"
theInstalledObjects(7) = "IISSample.PageCounter"
theInstalledObjects(8) = "MSWC.PermissionChecker"
theInstalledObjects(9) = "Scripting.FileSystemObject"
theInstalledObjects(10) = "adodb.connection"
theInstalledObjects(11) = "SoftArtisans.FileUp"
theInstalledObjects(12) = "SoftArtisans.FileManager"
theInstalledObjects(13) = "JMail.SMTPMail"
theInstalledObjects(14) = "CDONTS.NewMail"
theInstalledObjects(15) = "Persits.MailSender"
theInstalledObjects(16) = "LyfUpload.UploadFile"
theInstalledObjects(17) = "Persits.Upload.1"
'进行50万次计算
dim t1,t2,lsabc,thetime
t1=timer
for i=1 to 500000
	lsabc= 1 + 1
next
t2=timer
thetime=cstr(int(((t2-t1)*10000)+0.5)/10)
%>
<table cellpadding="0" cellspacing="1" class="border">
  <tr><th colspan="2">用户信息</th></tr>
  <tr class="tdbg">
    <td class="t-right h25">当前用户：</td>
    <td><%=Request.Cookies("admin")("user")%></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">姓名身份：</td>
    <td><%=Request.Cookies("admin")("name")%>（<%=Request.Cookies("admin")("title")%>）</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">用户操作：</td>
    <td>第<span class="red"><%=Request.Cookies("admin")("gonum")%></span>次登录&nbsp;(上次登录时间：<span class="red"><%=Request.Cookies("admin")("lasttime")%></span>)</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">登录追踪：</td>
    <td>本次IP：<font color="red"><%=Request.Cookies("admin")("ip")%></font>&nbsp;&nbsp;上次IP：<span class="red"><%=Request.Cookies("admin")("oldip")%></span></td>
  </tr>
  <tr><th colspan="2">系统信息</th></tr>
  <tr class="tdbg">
    <td class="t-right h25">服务器及域名：</td>
    <td><%=Request.ServerVariables("server_name")%> / <span class="red"><%=Request.ServerVariables("Http_HOST")%></span></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">服务器类型：</td>
    <td><%=Request.ServerVariables("OS")%>(IP:<span class="red"><%=Request.ServerVariables("LOCAL_ADDR")%></span>)</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">脚本超时时间：</td>
    <td><%=Server.ScriptTimeout%> 秒</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">Jmail组件支持：</td>
    <td><%If Not IsObjInstalled(theInstalledObjects(13)) Then%> <span class="red">×</span> <%else%> √ <%end if%></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">CDONTS组件支持：</td>
    <td><%If Not IsObjInstalled(theInstalledObjects(14)) Then%> <span class="red">×</span> <%else%> √ <%end if%></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">数据库的使用：</td>
    <td><%If Not IsObjInstalled(theInstalledObjects(10)) Then%> <span class="red">×</span> <%else%> √ <%end if%></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">FSO&nbsp;文本读写：</td>
    <td><%If Not IsObjInstalled(theInstalledObjects(9)) Then%> <span class="red">×</span> <%else%> √ <%end if%></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">运算速度测试：</td>
    <td>Script Execution Time:<span class="red"><%=thetime%></span>ms (进行<span class="red">500'000</span>次计算使用时间：<span class="red"><%=thetime%></span>毫秒)
	</td>
  </tr>
</table>
<%call CloseConn()%>
</body>
</html>