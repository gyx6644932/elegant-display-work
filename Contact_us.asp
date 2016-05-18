<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file="inc/conn.asp" -->
<!--#include file="inc/function.asp" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>elegant-display</title>
<meta name="keywords" content="elegant display">
<meta name="description" content="elegant display">
<link href="css/style.css" type="text/css" rel="stylesheet">
<link rel="shortcut icon" href="img/favicon.ico" type="img/x-icon">
<link rel="Bookmark" href="img/favicon.ico">
<style type="text/css">
#allmap {
	width: 965px;
	height: 400px;
	overflow: hidden;
	margin: 0;
}
#l-map {
	height: 400px;
	width: 100%;
	float: left;
	border-right: 2px solid #bcbcbc;
}
</style>
<script type="text/javascript" src="http://api.map.baidu.com/api?v=2.0&ak=A706d69c0ed0720ec513b4c83cec37e1"></script>
<script type="text/javascript" src="js/jquery.min.js"></script>
<script type="text/javascript" src="js/ddsmoothmenu.js"></script>
</head>
<body>
<!--#include file="header.asp"-->

<div id="contact_box">
  <div id="contact_box2">

    <div id="contact_txt">
      <h3>Contact Us</h3>
      <p>JIAXING ELEGANT DISPLAY CO., LTD</p>
      <p>Email：rita@elegant-display.com</p>
      <p>Contact Person:Rita ding</p> 
      <p>Mobile:+86 13758360692</p>
      <p>Tel:+86 573-86912085</p>
      <p>Fax:+86 573-86912086</p> 
      <p>Skype: ritading12</p>
      <p>website:www.elegant-display.com</p>
      <p>ADD：FOURTH FLOOR， NO.107, QINJIAN ROAD, HAIYAN, </p>
  <p>ZHEJIANG, CHINA</p>
      <div class="fill"></div>
    </div>
    
    <div class="contact-part">
      <div class="wpcf7" id="wpcf7-f62-w1-o1">
        <form action="/#wpcf7-f62-w1-o1" method="post" class="wpcf7-form" novalidate="novalidate">
          <div style="display: none;">
            <input type="hidden" name="_wpcf7" value="62">
            <input type="hidden" name="_wpcf7_version" value="3.5.4">
            <input type="hidden" name="_wpcf7_locale" value="en_US">
            <input type="hidden" name="_wpcf7_unit_tag" value="wpcf7-f62-w1-o1">
            <input type="hidden" name="_wpnonce" value="37ff97de01">
          </div>
          <p><span class="wpcf7-form-control-wrap name">
            <input type="text" name="name" value="" size="40" class="wpcf7-form-control wpcf7-text wpcf7-validates-as-required" aria-required="true" placeholder="Name">
            </span> </p>
          <p> <span class="wpcf7-form-control-wrap email">
            <input type="email" name="email" value="" size="40" class="wpcf7-form-control wpcf7-text wpcf7-email wpcf7-validates-as-required wpcf7-validates-as-email" aria-required="true" placeholder="Email">
            </span></p>
          <p><span class="wpcf7-form-control-wrap phone">
            <input type="tel" name="phone" value="" size="40" class="wpcf7-form-control wpcf7-text wpcf7-tel wpcf7-validates-as-required wpcf7-validates-as-tel" aria-required="true" placeholder="Phone">
            </span></p>
          <p><span class="wpcf7-form-control-wrap message">
            <textarea name="message" cols="30" rows="3" class="wpcf7-form-control wpcf7-textarea" placeholder="Message"></textarea>
            </span> </p>
          <p>
            <input type="submit" value="Submit" class="wpcf7-form-control wpcf7-submit">
   
            <img class="ajax-loader" src="http://brianzeng.me/wp-content/plugins/contact-form-7/images/ajax-loader.gif" alt="Sending ..." style="visibility: hidden;"></p>
          <div class="wpcf7-response-output wpcf7-display-none"></div>
        </form>
      </div>
    </div>
  </div>
  <div id="l-map"></div>
  <script type="text/javascript">

// 百度地图API功能
var map = new BMap.Map("l-map");            // 创建Map实例
map.centerAndZoom(new BMap.Point(120.757028,30.785786), 11);
var local = new BMap.LocalSearch(map, {
  renderOptions: {map: map, panel: "r-result"}
});
local.search("锦江商务楼");
</script> 
</div>
<!--#include file="footer.asp"-->
</body>
</html>
