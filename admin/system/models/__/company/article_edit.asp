<!--#include file="article_config.asp"-->
<%
'<><><><><>处理提交数据部分<><><><><>
dim edit_id
   edit_id=request("id")
if text.isNum(edit_id) then
   editStr=sys_str_edit
else
   editStr=sys_str_add
end if
   
if request.Form("edit")="ok" then
tf_title=request.form("tf_title")
tf_content=request.form("tf_content")
tf_note=request.form("tf_note")
tf_add_data=request.form("tf_add_data")
if isdate(tf_add_data)=false then tf_add_data=now()
tf_order_id=request.form("tf_order_id")

'<><><><><><>写入数据部分<><><><><><>
set rs=server.createobject("adodb.recordset") 
    if text.isNum(edit_id) then
	   exec="select * from "&db_table&" where id=" & edit_id  '判断，修改数据
       rs.open exec,conn,1,3
    else
       exec="select * from "&db_table                       '判断，添加数据
	   rs.open exec,conn,1,3
       rs.addnew
    end if
    if text.isNum(edit_id) and rs.eof then
	   response.Write(sys_tip_none)
    else
       rs("title") = tf_title
	   rs("content") = tf_content
	   rs("note") = tf_note
	   rs("add_data") = tf_add_data
	   if text.isNum(tf_order_id) then rs("order_id") = tf_order_id
	   '<><>判断是否正确写入<><>
	   if err<>0 then call alert.reFresh(editStr&sys_tip_false,"article_edit.asp?id="&edit_id)
	end if
	rs.update
	rs.close
set rs=nothing
    call alert.reFresh(editStr&sys_tip_ok,"article_manage.asp?typdB_id="&typdB_id&"&typdS_id="&typdS_id)
end if

'<><><><><><>读取数据部分<><><><><><>
id=request.QueryString("id") 
if text.isNum(id) then
   set rs=conn.execute("select * from "&db_table&" where id="&id) 
	   if not rs.eof then
	   tf_title=rs("title")
	   tf_content=rs("content")
	   tf_note=rs("note")
	   tf_add_data=rs("add_data")
	   tf_order_id=rs("order_id")
	   if text.isNum(tf_order_id)=false then tf_order_id=0
	   end if
	   rs.close
   set rs=nothing  
end if
%>
<body>
<TABLE border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<TR><td>

<!--#include file="../../models_system/articles/articles_edit_head.asp"-->

<table border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<form name="FormE" class="well form-vertical" method="post" action="">

<tr class="forumRow"><td align="right">标题：</td>
<td><input name="tf_title" type="text" id="tf_title" value="<%=tf_title%>"/></td>
</tr>

<%if db_types then '设置分类%>
<tr class="forumRow"><td align="right">分类：</td>
<td><!--#include file="../../models_system/articles/articles_type.asp"--></td>
</tr>
<%end if%>

<tr class="forumRow"><td align="right">内容：</td><td>
<%=articlesHtml.editBox("tf_content",tf_content,"","","")%>
</td></tr>

<tr class="forumRaw">
<td width="88" align="right" valign="top">
  <input name="id" type="hidden" value="<%=id%>" />
  <input name="edit" type="hidden" value="ok" />
</td><td><%=articlesHtml.submitBtn("",editStr)%></td></tr>
</form>
</table></td>
</tr></table>
</body>
</html>