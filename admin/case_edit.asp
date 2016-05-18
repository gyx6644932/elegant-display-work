<!--#include file="Chk.asp"-->
<%
'****************************************提交*********************************************
if Request.form("Oper") <> "" then
	id = go_num(request.form("id"),0)
	bid = split(request.form("bid"),"|#|")
	title = SafeChar(request.form("title"))
	img = request.form("img")
	if strLength(img) <= 0 then img = "/images/nopic.gif"
	if strLength(title) <= 0 then call jsouterr("网站名称不能为空！")
	set rs = server.createobject("adodb.recordset")
	sql = "select * from [case] where id = "&id
	rs.open sql,conn,1,3
	set sql = nothing
	if Request.form("Oper") = "addsave" then rs.addnew
	rs("title") = title
	rs("bid") = bid(0)
	rs("bname") = bid(1)
	rs("content") = SafeChar(request.form("content"))
	rs("home") = go_num(request.form("home"),0)
	rs("img") = img
	rs("url") = request.form("url")
	rs("addtime") = now()
	rs("px") = go_num(request.form("px"),99)
	rs.update
	rs.close
	set rs = nothing
	Response.Write("<script type='text/javascript'>window.parent.location.reload();</script>")
	response.End()
end if
id = go_num(request.QueryString("id"),0)
if id <= 0 then
	Oper = "addsave"
	bid = 0
	title = ""
	content = ""
	img = ""
	url = "http://"
	px = 99
	id = 0
	home = 0
else
	Oper = "edit"
	Set rs = server.CreateObject("adodb.recordset")
	sql = "select * from [case] where id = "&id
	rs.Open sql, conn, 1, 1
	title = rs("title")
	bid = rs("bid")
	content = rs("content")
	img = rs("img")
	url = rs("url")
	px = rs("px")
	home = rs("home")
	rs.close
	set rs = nothing
end if
%>
<link rel="stylesheet" href="/html/themes/default/default.css" />
<script charset="utf-8" src="/html/kindeditor-min.js"></script>
<script charset="utf-8" src="/html/lang/zh_CN.js"></script>
<script>
KindEditor.ready(function(K) {
	var editor = K.editor({allowFileManager : true});
	K('#Upimage1').click(function(){
		editor.loadPlugin('image', function() {
			editor.plugin.imageDialog({
				imageUrl : K('#img').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#img').val(url);
					editor.hideDialog();
				}
			});
		});
	});
});
</script>
<table cellpadding="0" cellspacing="1" class="border">
  <form method="post">
  <input name="id" type="hidden" value="<%=id%>" />
  <input type="hidden" name="Oper" value="<%=Oper%>" />
  <tr class="tdbg">
    <th style="width:12%">网站名称</th>
    <td><input type="text" name="title" value="<%=title%>" maxlength="50" style="width:100%;" /></td>
  </tr>
  <tr class="tdbg">
    <th>网站分类</th>
    <td><select name="bid" id="bid"><%
set rsb = server.CreateObject("adodb.recordset")
sql = "Select * from case_bid order by px asc"
rsb.Open sql,Conn,1,1
if not rsb.eof then
if bid = 0 then bid = rsb("id")
do while not rsb.eof
if Cstr(rsb("id")) = Cstr(bid) then bid = rsb("id")
%>
<option value="<%=rsb("id")%>|#|<%=rsb("title")%>" <%if Cstr(rsb("id")) = Cstr(bid) then Response.Write("selected='selected'")%>><%=rsb("title")%></option>
<%
rsb.movenext
loop
end if
rsb.close
set rsb=nothing
%></select></td>
  </tr>
  <tr>
    <th>网站介绍</th>
    <td><textarea name="content" style="width:100%;height:250px;"><%=content%></textarea></td>
  </tr>
  <tr class="tdbg">
		<th>缩略图片</th>
		<td><input type="text" name="img" id="img" style="width:70%" value="<%=img%>" />&nbsp;<input type="button" id="Upimage1" class="bt" value="上传图片" /></td>
	</tr>
	<tr class="tdbg">
		<th>网站地址</th>
		<td><input type="text" name="url" value="<%=url%>" style="width:100%;" /></td>
	</tr>
  <tr class="tdbg">
    <th>推　　荐</th>
    <td><input type="radio" name="home" value="1" <%if home = 1 then Response.Write("checked='checked'")%> /> 是　　<input type="radio" name="home" value="0" <%if home = 0 then Response.Write("checked='checked'")%> /> 否</td>
  </tr>
  <tr class="tdbg">
    <th>排　　序</th>
    <td><input type="text" name="px" value="<%=px%>" class="nochines" maxlength="8" style="width:60px;" /></td>
  </tr>
  <tr class="tdbg2">
    <td colspan="2"><input type="submit" class="bt" value="确认提交" /></td>
  </tr>
  </form>
</table>
<%call CloseConn()%>
</body>
</html>