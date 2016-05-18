<!--#include file="chk.asp"-->
<%
dim adminpower(100)
s=split(Request.Cookies("adminpower"),"|")
For i=0 to UBound(s)
	adminpower(i)=CBool(s(i))
Next
%>
<style type="text/css">
<!--
*{margin:0;padding:0;}
body {background-color:#0096CE;overflow-x:hidden;overflow-y:scroll;margin:32px 0 0 0;background-image:url(images/admin_title.gif);background-repeat:no-repeat;}
.menu_class{width:158px;}
.menu_class ul{list-style-type:none;overflow:hidden;cursor:pointer;height:25px;line-height:25px;}
.menu_class ul.open{height:auto;}
.menu_class ul span{display:block;background:url(images/Admin_left.gif) no-repeat;color:#276DBE;font-weight:bold;padding-left:5px;height:25px;line-height:25px;}
.menu_class ul.open span{display:block;background:url(images/Admin_left2.gif) no-repeat;color:#276DBE;font-weight:bold;padding-left:5px;height:25px;line-height:25px;}
.menu_class ul li{background-color:#FFFFFF;border-bottom:#0096CE 1px solid;height:25px;}
.menu_class ul li a{color:#276DBE;padding-left:10px;font-weight:normal;text-decoration:none;display:block;height:25px;}
.menu_class ul li a:hover,a.current{color:#276DBE;padding-left:10px;font-weight:normal;text-decoration:none;display:block;background:url(images/Admin_left3.gif) no-repeat;}
-->
</style>
<!--[if lt IE 8]>
<style type="text/css">
li a{display:inline-block;}
li a{display:block;}
</style>
<![endif]-->
<script type="text/javascript">
function menu_class(id,onlyone){
	if(!document.getElementById || !document.getElementsByTagName){return false;}
	this.menu=document.getElementById(id);
	this.submenu=this.menu.getElementsByTagName("ul");
	this.speed=3;
	this.time=10;
	this.onlyone=onlyone==true?onlyone:false;
	this.links = this.menu.getElementsByTagName("a");
}
menu_class.prototype.init=function(){
	var mainInstance = this;
	for(var i=0;i<this.submenu.length;i++){
		this.submenu[i].getElementsByTagName("span")[0].onclick=function(){
			mainInstance.toogleMenu(this.parentNode);
		};
	}
	for(var i=0;i<this.links.length;i++){
		this.links[i].onclick=function(){
			this.className = "current";
			mainInstance.removeCurrent(this);
		}
	}
}
menu_class.prototype.removeCurrent = function(link){
	for (var i = 0; i < this.links.length; i++){
		if (this.links[i] != link){
			this.links[i].className = " ";
		}
	}
}
menu_class.prototype.toogleMenu=function(submenu){
	if(submenu.className=="open"){
		this.closeMenu(submenu);
		}else{
		this.openMenu(submenu);
	}
}
menu_class.prototype.openMenu=function(submenu){
	var fullHeight=submenu.getElementsByTagName("span")[0].offsetHeight;
	var links = submenu.getElementsByTagName("a");
	for (var i = 0; i < links.length; i++){
		fullHeight += links[i].offsetHeight;
	}
	var moveBy = Math.round(this.speed * links.length);
	var mainInstance = this;
	var intId = setInterval(function(){
		var curHeight = submenu.offsetHeight;
		var newHeight = curHeight + moveBy;
		if (newHeight <fullHeight){
			submenu.style.height = newHeight + "px";
		}else{
			clearInterval(intId);
			submenu.style.height = "";
			submenu.className = "open";
		}
	}, this.time);
	this.collapseOthers(submenu);
}
menu_class.prototype.closeMenu=function(submenu){
	var minHeight=submenu.getElementsByTagName("span")[0].offsetHeight;
	var moveBy = Math.round(this.speed * submenu.getElementsByTagName("a").length);
	var mainInstance = this;
	var intId = setInterval(function(){
		var curHeight = submenu.offsetHeight;
		var newHeight = curHeight - moveBy;
		if (newHeight > minHeight){
			submenu.style.height = newHeight + "px";
		}else{
			clearInterval(intId);
			submenu.style.height = "";
			submenu.className = "";
		}
	}, this.time);
}
menu_class.prototype.collapseOthers = function(submenu){
	if(this.onlyone){
		for (var i = 0; i < this.submenu.length; i++){
			if (this.submenu[i] != submenu){
				this.closeMenu(this.submenu[i]);
			}
		}
	}
}
</script>
<div id="menu_class" class="menu_class" style="float:left;">
	<ul class="open">
		<span>���ù���</span>
		<li><a href="User_pass.asp?id=<%=Request.Cookies("admin")("id")%>" target="right">�޸ĵ�¼����</a></li>
		<li><a href="User_edit.asp?id=<%=Request.Cookies("admin")("id")%>" target="right">�޸ĸ�������</a></li>
	</ul>
	<%if adminpower(1) then%>
	<ul>
		<span>ϵͳ����</span>
		<%if adminpower(2) then%><li><a href="Admin.asp" target="right">��վ����Ա</a></li><%end if%>
		<%if adminpower(3) then%><li><a href="Backup.asp" target="right">���ݿ����</a></li><%end if%>
		<%if adminpower(4) then%><li><a href="UploadFileManage.asp" target="right">�ϴ��ļ�����</a></li><%end if%>
    </ul>
	<%end if%>
    <ul>
		<span>��վ����</span>
        <li><a href="xt_config.asp" target="right">��������</a></li>
        <li><a href="web.asp" target="right">��ҳ����</a></li>
        <li><a href="Banner.asp" target="right">�õƹ���</a></li>
        <li><a href="link.asp" target="right">��������</a></li>
		<li><a href="menu.asp" target="right">�����˵�</a></li>
        <li><a href="quickbar.asp" target="right">����ͨ��</a></li>
    </ul>
	<ul>
		<span>��վ�����ײ͹���</span>
        <li><a href="Web_list.asp" target="right">�ײ͹���</a></li>
    </ul>
    <ul>
		<span>��վ����</span>
        <li><a href="case_bid.asp" target="right">��������</a></li>
        <li><a href="case.asp" target="right">�����б�</a></li>
    </ul>
	<ul>
		<span>�ػݷ������</span>
        <li><a href="Web_fuwu.asp" target="right">�ػݷ������</a></li>
    </ul>
	<ul>
		<span>���Ź���</span>
        <li><a href="news_edit.asp" target="right">�������</a></li>
		<li><a href="news.asp" target="right">���Ź���</a></li>
    </ul>
	<ul>
		<span>����ͳ��</span>
		<li><a href="stat.asp" target="right">��������</a></li>
		<li><a href="tuiguang.asp" target="right">��վ�ƹ�</a></li>
	</ul>
	<ul>
		<span>��Ȩ��Ϣ</span>
		<li><a href="http://sighttp.qq.com/authd?IDKEY=81fcbd1fad200e4c1474ce2c2559660d63fa6c3362e7151a" target="_blank">����������Chu</a></li>
		<li><a href="http://www.diwei8.com/" target="_blank">�������</a></li>
	</ul>
</div>
<!-- 1 - 67 ,99 -->
<script type="text/javascript">
window.onload = function() {
myMenu = new menu_class("menu_class",true);
myMenu.init();
};
</script>
<%call CloseConn()%>
</body>
</html>