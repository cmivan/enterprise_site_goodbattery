<!--#include file="article_config.asp"-->
<!--#include file="../../libraries/articles.asp"-->
<%
articles.table = db_table


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
tf_type_id=request.Form("type_id")
tf_content=request.form("tf_content")
tf_note=request.form("tf_note")
tf_add_data=request.form("tf_add_data")
if isdate(tf_add_data)=false then tf_add_data=now()
tf_order_id=request.form("tf_order_id")

if db_types<>0 then
'——————分类
  '//系判断是否为数组，再判断是否为数字(否则失败)
  if tf_type_id<>"" then
	  tf_type_id = tf_type_id & ",0"
	  type_ids=split(tf_type_id,",")
	  if ubound(type_ids)>=1 then
		 if type_ids(0)<>"" and isnumeric(type_ids(0)) and type_ids(1)<>"" and isnumeric(type_ids(1)) then
			tf_typeB_id=type_ids(0)
			tf_typeS_id=type_ids(1)
			session("type_id") = tf_type_id   '记录本次操作分类
		 else
			alert.msgBack("分类有误!请重新选择.")
		 end if
	  end if
   end if
end if

'<><><><><><>写入数据部分<><><><><><>
articles.rsUpdate( edit_id )
    if text.isNum(edit_id) and articles.rs.eof then
	   response.Write( sys_tip_none )
    else
       articles.rs("title") = tf_title
	   articles.rs("typeB_id") = tf_typeB_id
	   articles.rs("typeS_id") = tf_typeS_id
	   articles.rs("content") = tf_content
	   articles.rs("note") = tf_note
	   articles.rs("add_data") = tf_add_data
	   if text.isNum(tf_order_id) then articles.rs("order_id") = tf_order_id
	   if err<>0 then call alert.reFresh(editStr&sys_tip_false,"?id="&edit_id)
	end if
articles.rsUpdateClose()

	
    call alert.reFresh(editStr&sys_tip_ok,"article_manage.asp?typdB_id="&typdB_id&"&typdS_id="&typdS_id)
end if

'<><><><><><>读取数据部分<><><><><><>
id=request.QueryString("id") 
if text.isNum(id) then
   set rs=conn.execute("select * from "&db_table&" where id="&id) 
	   if not rs.eof then
	   tf_title=rs("title")
	   
	   tf_typeB_id =rs("typeB_id")
	   tf_typeS_id =rs("typeS_id")
	   '记录分类,用于分类下拉
	   if tf_typeS_id="" or isnumeric(tf_typeS_id)=false then
	      session("type_id")=tf_typeB_id
	   else
	      session("type_id")=tf_typeB_id&","&tf_typeS_id
	   end if
		  
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