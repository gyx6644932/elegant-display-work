<!--#include file="chk.asp"-->
<!--#include file="Upload_top.asp"-->
<table cellpadding="0" cellspacing="1" class="border">
  <tr>
    <th>文件夹名称</th>
    <th>类型</th>
    <th>管理</th>
  </tr>
  <tr class="tdbg">
    <td class="t-center h25">Image</td>
    <td class="t-center">图片(gif、jpg、jpeg、png、bmp)</td>
    <td class="t-center"><a href="Upload_image.asp">进入管理</a></td>
  </tr>
  <tr class="tdbg">
    <td class="t-center h25">Flash</td>
    <td class="t-center">动画(swf)</td>
    <td class="t-center"><a href="Upload_flash.asp">进入管理</a></td>
  </tr>
  <tr class="tdbg">
    <td class="t-center h25">Media</td>
    <td class="t-center">媒体(flv、mp3、wav、wma、wmv、mid、avi、mpg、asf、rm、rmvb)</td>
    <td class="t-center"><a href="Upload_media.asp">进入管理</a></td>
  </tr>
  <tr class="tdbg">
    <td class="t-center h25">File</td>
    <td class="t-center">附件(doc、docx、xls、xlsx、ppt、txt、zip、rar、gz、bz2、pdf)</td>
    <td class="t-center"><a href="Upload_file.asp">进入管理</a></td>
  </tr>
  <tr class="tdbg2">
    <td colspan="3"><span class="red">（请定期清理无用的文件，以免影响服务器性能，导致降低网站运行速度！）</span></td>
  </tr>
</table>
<%call CloseConn()%>
</body>
</html>