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
'����50��μ���
dim t1,t2,lsabc,thetime
t1=timer
for i=1 to 500000
	lsabc= 1 + 1
next
t2=timer
thetime=cstr(int(((t2-t1)*10000)+0.5)/10)
%>
<table cellpadding="0" cellspacing="1" class="border">
  <tr><th colspan="2">�û���Ϣ</th></tr>
  <tr class="tdbg">
    <td class="t-right h25">��ǰ�û���</td>
    <td><%=Request.Cookies("admin")("user")%></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">������ݣ�</td>
    <td><%=Request.Cookies("admin")("name")%>��<%=Request.Cookies("admin")("title")%>��</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�û�������</td>
    <td>��<span class="red"><%=Request.Cookies("admin")("gonum")%></span>�ε�¼&nbsp;(�ϴε�¼ʱ�䣺<span class="red"><%=Request.Cookies("admin")("lasttime")%></span>)</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">��¼׷�٣�</td>
    <td>����IP��<font color="red"><%=Request.Cookies("admin")("ip")%></font>&nbsp;&nbsp;�ϴ�IP��<span class="red"><%=Request.Cookies("admin")("oldip")%></span></td>
  </tr>
  <tr><th colspan="2">ϵͳ��Ϣ</th></tr>
  <tr class="tdbg">
    <td class="t-right h25">��������������</td>
    <td><%=Request.ServerVariables("server_name")%> / <span class="red"><%=Request.ServerVariables("Http_HOST")%></span></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">���������ͣ�</td>
    <td><%=Request.ServerVariables("OS")%>(IP:<span class="red"><%=Request.ServerVariables("LOCAL_ADDR")%></span>)</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�ű���ʱʱ�䣺</td>
    <td><%=Server.ScriptTimeout%> ��</td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">Jmail���֧�֣�</td>
    <td><%If Not IsObjInstalled(theInstalledObjects(13)) Then%> <span class="red">��</span> <%else%> �� <%end if%></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">CDONTS���֧�֣�</td>
    <td><%If Not IsObjInstalled(theInstalledObjects(14)) Then%> <span class="red">��</span> <%else%> �� <%end if%></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">���ݿ��ʹ�ã�</td>
    <td><%If Not IsObjInstalled(theInstalledObjects(10)) Then%> <span class="red">��</span> <%else%> �� <%end if%></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">FSO&nbsp;�ı���д��</td>
    <td><%If Not IsObjInstalled(theInstalledObjects(9)) Then%> <span class="red">��</span> <%else%> �� <%end if%></td>
  </tr>
  <tr class="tdbg">
    <td class="t-right h25">�����ٶȲ��ԣ�</td>
    <td>Script Execution Time:<span class="red"><%=thetime%></span>ms (����<span class="red">500'000</span>�μ���ʹ��ʱ�䣺<span class="red"><%=thetime%></span>����)
	</td>
  </tr>
</table>
<%call CloseConn()%>
</body>
</html>