<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Response.Cookies("CookieCheck")="on"%>
<head>
<title>网站后台登录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/login.css" rel="stylesheet" type="text/css" />
<Script Language="JavaScript">
	<!--
	function chk_data(){
		if (document.form1.username.value==""){
			alert("\操作出错，下面是产生错误的可能原因：\n\n·管理员名称不能为空！");
			document.form1.username.focus();
			return false;
		}
		if (document.form1.userpsw.value==""){
			alert("\操作出错，下面是产生错误的可能原因：\n\n·管理员密码不能为空！");
			document.form1.userpsw.focus();
			return false;
		}
		if (document.form1.checkcode.value==""){
			alert("\操作出错，下面是产生错误的可能原因：\n\n·请输入您的验证码！");
			document.form1.checkcode.focus();
			return false;
		}
		return true;
	}
	// -->
</Script>
<SCRIPT LANGUAGE="JavaScript">   
function AddFavorite(sURL, sTitle) {   
    try {   
        window.external.addFavorite(sURL, sTitle);   
    } catch (e) {   
        try {   
            window.sidebar.addPanel(sTitle, sURL, "");   
        } catch (e) {   
            alert("加入收藏失败，请使用Ctrl+D进行添加");   
        }   
    }   
}
</SCRIPT>
</head>
<body onLoad="document.form1.username.focus();">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr height="70px">
    <td></td>
  </tr>
  <tr>
    <td><div class="con">
        <div class="login">
          <div class="b_left"></div>
          <div class="input">
			<form method=POST action="login_chk.asp" name="form1" onSubmit="JavaScript: return chk_data();"><input type="hidden" name="cookieexists" value="false" readonly>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" class="logTb">
              <tr>
                <td colspan="3" style="padding-top:14px;"><img src="images/logo.jpg"></td>
                <td align="right" valign="top" style="padding-top:30px; padding-right:2px;"><a href="javascript:void(null)" onClick="AddFavorite(window.location,document.title)" title="加入浏览器收藏夹，方便下次登录" class="fav"></a></td>
              </tr>
              <tr>
                <th>用户名</th>
                <td align="right" colspan="2"><input type="text" name="username" maxlength="20" class="input-border" /></td>
                <td></td>
              </tr>
              <tr>
                <th>密&nbsp;&nbsp;&nbsp;&nbsp;码</th>
                <td align="left" colspan="2"><input type="password" name="userpsw" maxlength="20" class="input-border" /></td>
                <td></td>
              </tr>
              <tr>
                <th>验证码</th>
                <td align="left" width="100"><input type="text" name="checkcode" maxlength="6" class="input-code" style="ime-mode:disabled;" /></td>
                <td width="85"><img src="../inc/Code.asp?" onClick="this.src+=Math.random()" alt="图片看不清？点击重新得到验证码" style="cursor:hand;"></td>
                <td></td>
              </tr>
              <tr>
                <th></th>
                <td colspan="2"><input type="submit" value=" " class="login-b"  onMouseOver="this.className='login-b2'" onMouseDown="this.className='login-b3'" onMouseOut="this.className='login-b'"/></td>
              </tr>
              <tr>
                <td colspan="4" ><div class="reg">如果您不是本站管理员请离开或<a href="../">返回首页</a>&nbsp;！谢谢合作！</div></td>
              </tr>
            </table>
			</form>
          </div>
          <div class="b_right"></div>
        </div>
      </div></td>
  </tr>
</table>
<script language="JavaScript">
<!--
if (document.cookie.search("CookieCheck=") == -1) {
    alert("您的浏览器“Cookies”被关闭，请开启后再使用本站！");
	document.form1.cookieexists.value ="false";
	} else {
    document.form1.cookieexists.value ="true";
}
// -->
</script>
</body>
</html>