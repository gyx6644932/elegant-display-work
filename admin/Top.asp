<!--#include file="chk.asp"-->
<style type="text/css">
<!--
body {margin:0px;}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
function preloadImg(src)
{
	var img=new Image();
	img.src=src
}
preloadImg("Images/admin_top_open.gif");

var displayBar=true;
function switchBar(obj)
{
	if (displayBar)
	{
		parent.frame.cols="0,*";
		displayBar=false;
		obj.innerText="打开管理菜单 >>";
	}
	else{
		parent.frame.cols="178,*";
		displayBar=true;
		obj.innerText="<< 隐藏管理菜单";
	}
}
</script>
<table width="100%" height="35"  border="0" cellpadding="0" cellspacing="0" background="img/admin_top.jpg" bgcolor="#0096CE" class="tb_1">
  <tr>
    <td width="284">&nbsp;&nbsp;<a href="javascript:void(null)" onClick="switchBar(this)" class="atop">&lt;&lt;&nbsp;隐藏管理菜单</a></td>
    <td width="490">
<font color="#FFFFFF">
<%=formatdatetime(now(),1)%>&nbsp;&nbsp;<%=weekdayname(weekday(now))%>
<SPAN id=liveclock 15px? height: 109px; style?="width:"></SPAN>
<SCRIPT language=javascript>
function www_helpor_net()
{
var Digital=new Date()
var hours=Digital.getHours()
var minutes=Digital.getMinutes()
var seconds=Digital.getSeconds() 
if(minutes<=9)
minutes="0"+minutes
if(seconds<=9)
seconds="0"+seconds
myclock=hours+":"+minutes+":"+seconds
if(document.layers){document.layers.liveclock.document.write(myclock)
document.layers.liveclock.document.close()
}else if(document.all)
liveclock.innerHTML=myclock
setTimeout("www_helpor_net()",1000)
}
www_helpor_net();
//-->
</SCRIPT></font>
</font>
	</td>
    <td width="235" ><a href="../" target="_blank" class="atop">浏览前台</a>&nbsp;&nbsp;<a href="../admin" target="_top" class="atop">后台首页</a>&nbsp;&nbsp;<a href="logout.asp" target="_top" onClick="return confirm('你真的要退出后台吗？')" class="atop">安全退出</a> </td>
  </tr>
</table>