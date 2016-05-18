<%option explicit%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file="inc/conn.asp" -->
<!--#include file="inc/function.asp" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="keywords" content="elegant display">
<meta name="description" content="elegant display">
<link href="css/style.css" type="text/css" rel="stylesheet">
<link rel="shortcut icon" href="img/favicon.ico" type="img/x-icon">
<link rel="Bookmark" href="img/favicon.ico">
<script type="text/javascript" src="js/jquery.min.js"></script>
<script type="text/javascript" src="js/ddsmoothmenu.js"></script>
<title>elegant-display</title>
</head>
<body>
<!--#include file="header.asp"-->
<!--Top End-->
<div class="banner-box">
  <div class="slide-box">
    <div class="slide-item" style="display: none; "><a href="#"><img src="images/sy1.jpg"></a></div>
    <div class="slide-item" style="display: block; "><a href="#"><img src="images/sy2.jpg"></a></div>
    <div class="slide-item" style="display: none; "><a href="#"><img src="images/sy3.jpg"></a></div>
    <div class="slide-item" style="display: none; "><a href="#"><img src="images/sy4.jpg"></a></div>
    <div class="slide-item" style="display: none; "><a href="#"><img src="images/sy5.jpg"></a></div>
    <div class="slide-box-masker" style="cursor: pointer; position: absolute; top: 0px; left: 0px; background-color: rgb(0, 0, 0); width: 100%; height: 100%; display: none; opacity: 1; background-position: initial initial; background-repeat: initial initial; "></div>
  </div>
  <div class="slide-snap-box">
    <ul>
      <li class=""> <img src="images/55.png">
        <div>
          <p class="slide-snap-item-title">Elegant Display</p>
          <p class="slide-snap-item-intro">we have grown rapidly and now employ</p>
        </div>
      </li>
      <li class="slide-snap-item-current"> <img src="images/33.jpg">
        <div>
          <p class="slide-snap-item-title">Fashion Trend</p>
          <p class="slide-snap-item-intro">We work closely with retail design teams </p>
        </div>
      </li>
      <li class=""> <img src="images/66.jpg">
        <div>
          <p class="slide-snap-item-title">High Quality</p>
          <p class="slide-snap-item-intro">We provide custom-made for clients. Quality, design and service</p>
        </div>
      </li>
      <li class=""> <img src="images/11.png">
        <div>
          <p class="slide-snap-item-title">European Standard</p>
          <p class="slide-snap-item-intro">The products are mainly exported to USA, Germany, France, Japan, South Africa etc. </p>
        </div>
      </li>
      <li> <img src="images/22.jpg">
        <div>
          <p class="slide-snap-item-title">Customize Models</p>
          <p class="slide-snap-item-intro">Welcome customers from all over the wold to visit our factory!</p>
        </div>
      </li>
    </ul>
  </div>
</div>

<script src="js/index.js" type="text/javascript"></script>
<div class="clear"></div>



<div class="gdtp">

  <div class="index_pro_list">
    <div class="pro_list_left"><a onmouseup="javascript:spd(30)" onmousedown="javascript:spd(5)" onmouseover="javascript:left()" style="cursor:hand"></a></div>
    <div id="demo">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tbody>
          <tr>
            <td id="marquePic1"><div class="pro_list_center">
              <%
			 
						dim rsc : set rsc = server.createobject("adodb.recordset")
						sql = "select top 30 * from [Case] where home = 1 order by px asc,id desc" 
						rsc.open sql, conn, 1, 1
						set sql = nothing
						if not rsc.eof then
						do while not rsc.eof
						
						
	            %>
              <a href="<%=rsc("url")%>" target="_blank"><img src="<%=rsc("img")%>" height="140" border="0"></a>
              <%
			   rsc.movenext : loop
			   end if
			   	rsc.close
	            set rsc = nothing
			    %>
            <td id="marquePic2"><div class="pro_list_center">
                <%
			 
						dim rsd : set rsd = server.createobject("adodb.recordset")
						sql = "select top 30 * from [Case] where home = 1 order by px asc,id desc" 
						rsd.open sql, conn, 1, 1
						set sql = nothing
						if not rsd.eof then
						do while not rsd.eof
						
						
	            %>
                <a href="<%=rsd("url")%>" target="_blank"><img src="<%=rsd("img")%>" height="140" border="0"></a>
                <%
			   rsd.movenext : loop
			   end if
			   	rsd.close
	            set rsd = nothing
			    %>
              </div></td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>
  <div class="pro_list_right"><a onmouseup="javascript:spd(30)" onmousedown="javascript:spd(5)" onmouseover="javascript:right()" style="cursor:hand"></a></div>
</div>
<script type="text/javascript"> 
var istop=1;
var LorR=0;
var speed=30; 
var demo = document.getElementById("demo");
var marquePic1 = document.getElementById("marquePic1");
var marquePic2 = document.getElementById("marquePic2");
marquePic2.innerHTML=marquePic1.innerHTML; 
demo.scrollLeft=marquePic1.scrollWidth/2;
function left(){
	LorR=0;
}
function right(){
	LorR=1;
}
function spd(n){
	speed=n;
	clearInterval(MyMar);
	MyMar=setInterval(Marquee,speed); 
}	
function Marquee(){
if(LorR) MarqueeRight();
else MarqueeLeft();
}
function MarqueeLeft(){
	with (demo) {
        if(scrollLeft>=marquePic1.scrollWidth){ 
            scrollLeft=0
        }else{ 
            scrollLeft++
        }
    }
}
function MarqueeRight(){
	with (demo) {
        if(scrollLeft<=0){ 
            scrollLeft=marquePic1.scrollWidth 
        }else{ 
            scrollLeft--
        }
    }
} 
var MyMar=setInterval(Marquee,speed) 
demo.onmouseover=function() {clearInterval(MyMar)} 
demo.onmouseout=function() {MyMar=setInterval(Marquee,speed)} 
</script>
</div>
<!--#include file="footer.asp"--> 
</body>
</html>
