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

<script type="text/javascript" src="js/jquery.min.js"></script>
<script type="text/javascript" src="js/ddsmoothmenu.js"></script>
</head>
<body>
<!--#include file="header.asp"-->
<div id="cpdt"> <img src="images/cpdt.jpg" width="956" height="230"> </div>
<div class="mod">
	<div class="product_left left">
    <div class="hd1"></div>
	
     <div id="sidebar" class="four columns">
	<ul class="sidenav">
	<li><a href="cpjs.asp">Female Mannequins</a>
<ul class="children">
	 <li class="page_item page-item-2601"><a href="cpjs.asp"><em>01&nbsp;&nbsp;</em>Combination</a></li>
            <li class="page_item page-item-2610"><a href="cpjs.asp"><em>02&nbsp;&nbsp;</em>Abstract</a></li>
            <li class="page_item page-item-2621"><a href="cpjs.asp"><em>03&nbsp;&nbsp;</em>Realistic </a></li>
            <li class="page_item page-item-2621"><a href="cpjs.asp"><em>04&nbsp;&nbsp;</em>Headless </a></li>
</ul>
</li>
<li class="page_item page-item-2565"><a href="cpjs2.asp">Male Mannequins </a>
<ul class="children">
	<li class="page_item page-item-2656"><a href="cpjs2.asp"><em>01&nbsp;&nbsp;</em>Abstract</a></li>
	<li class="page_item page-item-2662"><a href="cpjs2.asp"><em>02&nbsp;&nbsp;</em>Headles</a></li>
</ul>
</li>
<li class="page_item page-item-2474"><a href="cpjs3.asp">Children Mannequins </a></li>
<li class="current_page_item"><a href="cpjs4.asp">Torso </a></li>
</ul>
	
	

<!-- begin generated sidebar -->

<!-- end generated sidebar -->

	
	</div>
    
    
    
    
	</div>
	<div class="product_right right">
		<div class="product_path">
        <div class="path2">
        <a href="../">HOME</a> &gt;Children Mannequins 
        
        </div>
        </div>
		<div class="clear"></div>
		<div class="news_detail">
			
			<div class="clear"></div>
			



<div class="news_nr">
	
			
			<div class="caseshow">
				<%
                Set rs = server.CreateObject("adodb.recordset")
                sql = "select * from [Case4]"
				
				sql = sql&" order by px asc,id desc"
                rs.Open sql, conn, 1, 1
                If not (rs.bof and rs.EOF) Then
				rs.PageSize = 6
				iCount = rs.RecordCount 
				iPageSize = rs.PageSize
				maxpage = rs.PageCount
				page = request("page")
				If Not IsNumeric(page) Or page = "" Then
					page = 1
				Else
					page = CInt(page)
				End If
				If page<1 Then
					page = 1
				ElseIf page>maxpage Then
					page = maxpage
				End If
				rs.AbsolutePage = Page
				If page = maxpage Then
					x = iCount - (maxpage -1) * iPageSize
				Else
					x = iPageSize
				End If
				For i = 1 To x
				
				
                %>
				
				<li><img src="../<%=rs("img")%>" /></a>
					
				</li>
				<%
                rs.movenext
				Next
				End If
                rs.close
                set rs = nothing
                %>
				<li style="height:30px;line-height:30px;">
					<%Call PageControl(iCount, maxpage, page, iPageSize)%>
				</li>
			</div>
			<div id="clear">&nbsp;</div>
</div>


		</div>
		<div class="clear"></div>
	</div>
	<div class="clear"></div>
</div>

<!--#include file="footer.asp"-->
</body>
</html>
