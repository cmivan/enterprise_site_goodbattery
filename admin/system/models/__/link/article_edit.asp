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
tf_small_pic=request.form("tf_small_pic")
tf_toUrl=request.form("tf_toUrl")
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
	   articles.rs("small_pic") = tf_small_pic
	   articles.rs("toUrl") = tf_toUrl
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
	   tf_small_pic=rs("small_pic")  
	   tf_toUrl=rs("toUrl")
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
<tr><td>

<!--#include file="../../models_system/articles/articles_edit_head.asp"-->

<table border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<form name="article_update" method="post" action="">
<tr class="forumRow">
<td width="13%" align="right">名称：</td>
<td width="87%">
<input name="tf_title" type="text" id="tf_title" value="<%=tf_title%>" size="40" /></td></tr>
	  
<%
'设置分类
if db_types<>0 then
%>
<tr class="forumRow">
<td align="right">分类：</td>
<td>
<!--#include file="../../models_system/articles/articles_type.asp"--></td>
</tr>
<%end if%>
		  
		  
<tr class="forumRow">
<td align="right" class="forumRow">图标：</td>
<td>
<input name="tf_small_pic" type="text" id="tf_small_pic" value="<%=tf_small_pic%>" size="45" maxlength="255" />
<%=articlesHtml.upImgBtn("上传图片","article_update.small_pic")%>
</td></tr>
		  
<tr class="forumRow">
<td align="right"><%=db_title%>地址：</td>
<td><input name="tf_toUrl" type="text" id="tf_toUrl" value="<%=tf_toUrl%>" size="45" maxlength="255" /> 
(<span class="red">链接地址如: http://www.baidu.com</span>) </td>
</tr>
		  
<tr class="forumRow">
<td align="right" valign="top" style="padding-top:4px;">简要：</td>
<td><%=articlesHtml.noteBox("tf_note",tf_note)%></td>
</tr>
		  
<tr style="display:none" class="forumRaw">
<td align="right" valign="top">排序：</td>
<td><input name="tf_order_id" type="text" id="tf_order_id" value="<%=tf_order_id%>" size="10" /></td>
</tr>

<tr class="forumRaw">
<td align="right" valign="top" class="forumRow"><p class="submit">
<input name="id" type="hidden" value="<%=id%>" />
<input name="edit" type="hidden" value="ok" />
</p></td>
<td><span class="submit">
<%=articlesHtml.submitBtn("",editStr)%>
</span></td></tr>
</form></table>
</td></TR></TABLE>
</body>
</html>