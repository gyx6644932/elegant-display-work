<!--#include file="chk.asp"-->
<%
if Request.form("Oper") <> "" then
	title = ReplaceBadChar(request.form("title"))
	imgpath = request.Form("imgpath")
	link = request.form("link")
	id = ReplaceBadChar(request.form("id"))
	px = ReplaceBadChar(request.form("px"))
	IF strLength(id) <= 0 or not isshuzi(id) then jsouterr("�����������")
	IF strLength(px) <= 0 or not isshuzi(px) then jsouterr("�������Ϊ���֣��Ҳ���Ϊ�գ�")
	IF strLength(imgpath) <= 0 then jsouterr("������ͼƬ��ַ���ϴ�ͼƬ��")
	set rs = server.createobject("adodb.recordset")
	sql = "select * from flash where id = "&id
	rs.open sql,conn,1,3
	if Request.form("Oper") = "addsave" then rs.addnew
	rs("imgpath") = imgpath
	rs("title") = title
	rs("px") = px
	rs("link") = link
	rs.update
	rs.close
	set rs = nothing
	Response.Write("<script type='text/javascript'>window.parent.location.reload();</script>")
	response.End()
end if
'����id���������޸ģ�������������
id = request.QueryString("id")
if id <> "" and isnumeric(id) then
	Oper = "edit"
	Set rs = server.CreateObject("adodb.recordset")
	sql = "select * from flash where id = "&id
	rs.Open sql, conn, 1, 1
	px = rs("px")
	title = rs("title")
	imgpath = rs("imgpath")
	link = rs("link")
	rs.close
	set rs = nothing
else
	Oper = "addsave"
	link = "http://"
	px = 99
	id = 0
end if
%>
<script charset="utf-8" src="../html/kindeditor-min.js"></script>
<script charset="utf-8" src="../html/lang/zh_CN.js"></script>
<link rel="stylesheet" href="../html/themes/default/default.css" />
<script>
KindEditor.ready(function(K) {
	var editor = K.editor({allowFileManager : true});
	K('#Upimage').click(function(){
		editor.loadPlugin('image', function() {
			editor.plugin.imageDialog({
				imageUrl : K('#imgpath').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#imgpath').val(url);
					editor.hideDialog();
				}
			});
		});
	});
});
</script>
<table cellpadding="0" cellspacing="1" class="border">
  <form method="post">
  <input type="hidden" name="Oper" value="<%=Oper%>" />
  <input type="hidden" name="id" value="<%=id%>" />
  <tr class="tdbg">
    <th class="w15">��������</th>
    <td><input name="title" type="text" class="w100" value="<%=title%>" /></td>
  </tr>
  <tr class="tdbg">
    <th>����ͼƬ</th>
    <td><input name="imgpath" id="imgpath" type="text" class="w50" value="<%=imgpath%>" maxlength="240" />&nbsp;<input type="button" id="Upimage" class="bt" value="�ϴ�ͼƬ" /></td>
  </tr>
  <tr class="tdbg">
    <th>���ӵ�ַ</th>
    <td><input name="link" type="text" class="w100" value="<%=link%>" /></td>
  </tr>
  <tr class="tdbg">
    <th>�š�����</th>
    <td style="text-align:left;"><input name="px" type="text" value="<%=px%>" class="w100 nochines" /></td>
  </tr>
  <tr class="tdbg2">
    <td colspan="2"><input type="submit" class="bt" value="ȷ���ύ" /></td>
  </tr>
  </form>
</table>