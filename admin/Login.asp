<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Response.Cookies("CookieCheck")="on"%>
<head>
<title>��վ��̨��¼</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/login.css" rel="stylesheet" type="text/css" />
<Script Language="JavaScript">
	<!--
	function chk_data(){
		if (document.form1.username.value==""){
			alert("\�������������ǲ�������Ŀ���ԭ��\n\n������Ա���Ʋ���Ϊ�գ�");
			document.form1.username.focus();
			return false;
		}
		if (document.form1.userpsw.value==""){
			alert("\�������������ǲ�������Ŀ���ԭ��\n\n������Ա���벻��Ϊ�գ�");
			document.form1.userpsw.focus();
			return false;
		}
		if (document.form1.checkcode.value==""){
			alert("\�������������ǲ�������Ŀ���ԭ��\n\n��������������֤�룡");
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
            alert("�����ղ�ʧ�ܣ���ʹ��Ctrl+D�������");   
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
                <td align="right" valign="top" style="padding-top:30px; padding-right:2px;"><a href="javascript:void(null)" onClick="AddFavorite(window.location,document.title)" title="����������ղؼУ������´ε�¼" class="fav"></a></td>
              </tr>
              <tr>
                <th>�û���</th>
                <td align="right" colspan="2"><input type="text" name="username" maxlength="20" class="input-border" /></td>
                <td></td>
              </tr>
              <tr>
                <th>��&nbsp;&nbsp;&nbsp;&nbsp;��</th>
                <td align="left" colspan="2"><input type="password" name="userpsw" maxlength="20" class="input-border" /></td>
                <td></td>
              </tr>
              <tr>
                <th>��֤��</th>
                <td align="left" width="100"><input type="text" name="checkcode" maxlength="6" class="input-code" style="ime-mode:disabled;" /></td>
                <td width="85"><img src="../inc/Code.asp?" onClick="this.src+=Math.random()" alt="ͼƬ�����壿������µõ���֤��" style="cursor:hand;"></td>
                <td></td>
              </tr>
              <tr>
                <th></th>
                <td colspan="2"><input type="submit" value=" " class="login-b"  onMouseOver="this.className='login-b2'" onMouseDown="this.className='login-b3'" onMouseOut="this.className='login-b'"/></td>
              </tr>
              <tr>
                <td colspan="4" ><div class="reg">��������Ǳ�վ����Ա���뿪��<a href="../">������ҳ</a>&nbsp;��лл������</div></td>
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
    alert("�����������Cookies�����رգ��뿪������ʹ�ñ�վ��");
	document.form1.cookieexists.value ="false";
	} else {
    document.form1.cookieexists.value ="true";
}
// -->
</script>
</body>
</html>