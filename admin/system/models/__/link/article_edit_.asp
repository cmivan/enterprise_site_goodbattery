<!--#include file="article_config.asp"-->

<%
'///////////  处理提交数据部分 ////////// 
   edit_id   =request("id")
if edit_id="" or isnumeric(edit_id)=false then
   editStr="添加"
   else
   editStr="修改"
end if
   
if request.Form("edit")="ok" then
'///////////  写入数据部分 //////////
   set rs=server.createobject("adodb.recordset") 
    if edit_id="" or isnumeric(edit_id)=false then
       exec="select * from "&db_table                       '判断，添加数据
	   rs.open exec,conn,1,3
       rs.addnew
	else
	   exec="select * from "&db_table&" where id="&edit_id  '判断，修改数据
       rs.open exec,conn,1,3
	end if

	if edit_id<>"" and isnumeric(edit_id) and rs.eof then
	   response.Write("写入数据失败!")
	else

'——————接收数据
    title     =request.Form("title")
    content   =request.Form("content")
    type_id   =request.Form("type_id")
    big_pic   =request.Form("big_pic")
    small_pic =request.Form("small_pic")
    order_id  =request.Form("order_id")
	
	strong    =request.Form("strong")
	color     =request.Form("color")
	sizes     =request.Form("size")
	toUrl     =request.Form("toUrl")
	note      =request.Form("note")
	add_data  =request.Form("add_data")
 if isdate(add_data)=false then add_data=now()
	

'——————审核\热门\最新\推荐
	ok  =request.Form("ok")
	hot =request.Form("hot")
	news=request.Form("news")
	tj  =request.Form("tj")


'——————排序
 if order_id="" or isnumeric(order_id)=false then order_id=0







if db_types<>0 then
'////////////////////////////////////////////////////////////////////////////
'——————分类
  '//系判断是否为数组，再判断是否为数字(否则失败)
  type_ids=split(type_id,",")
  if ubound(type_ids)=1 then
     if type_ids(0)<>"" and isnumeric(type_ids(0)) and type_ids(1)<>"" and isnumeric(type_ids(1)) then
	    typeB_id=type_ids(0)
	    typeS_id=type_ids(1)
	    session("type_id")=type_id   '记录本次操作分类
     else
	    response.Write("<script>alert('分类有误!请重新选择.');history.back(1);</script>")
	    response.End()
     end if
  else
     if type_id<>"" and isnumeric(type_id) then
	    typeB_id=type_id
	    typeS_id=null
	    session("type_id")=type_id   '记录本次操作分类
	 else
	    response.Write("<script>alert('分类有误!请重新选择.');history.back(1);</script>")
	    response.End()
	 end if
  end if
'////////////////////////////////////////////////////////////////////////////
end if



	rs("title")    =title
	rs("content")  =content
	rs("typeB_id") =typeB_id
	rs("typeS_id") =typeS_id
	
	rs("big_pic")  =big_pic
	rs("small_pic")=small_pic
	rs("order_id") =order_id
	
	rs("toUrl") =toUrl
	rs("note")  =note
	rs("add_data")=add_data
	

    response.Write("<script>alert('"&editStr&"操作成功!');window.location.href='article_manage.asp?typdB_id="&typdB_id&"&typdS_id="&typdS_id&"';</script>")
	end if
	rs.update
	rs.close
set rs=nothing

end if





'#################  读取数据部分 #################

   id=request.QueryString("id") 
if id<>"" and isnumeric(id) then
  '/// 当前状态:编辑
   edit_stat="edit"
   set rs=server.createobject("adodb.recordset") 
       exec="select * from "&db_table&" where id="&id
       rs.open exec,conn,1,1 
	   if not rs.eof then
	
	      title     =rs("title")
	      content   =rs("content")
		  typeB_id  =rs("typeB_id")
	      typeS_id  =rs("typeS_id")
		  
		  '记录分类,用于分类下拉
		  if typeS_id="" or isnumeric(typeS_id)=false then
		     session("type_id")=typeB_id
		  else
		     session("type_id")=typeB_id&","&typeS_id
		  end if


	      big_pic   =rs("big_pic")
	      small_pic =rs("small_pic")
	      order_id  =rs("order_id")

		  strong    =rs("strong")
		  color     =rs("color")
		  sizes     =rs("size")

		  toUrl=rs("toUrl")
		  note =rs("note")
		  add_data=rs("add_data")

		  if order_id="" or isnumeric(order_id)=false then order_id=0

	   end if
	   rs.close
   set rs=nothing
   
else
  '/// 当前状态:添加
   edit_stat="add"
end if


  '/// 处理添加时间
  if isdate(add_data)=false then add_data=now()
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
<input name="title" type="text" value="<%=title%>" size="40" /></td></tr>
	  
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
<input name="small_pic" type="text" id="small_pic" value="<%=small_pic%>" size="45" maxlength="255" />
<%=articlesHtml.upImgBtn("上传图片","article_update.small_pic")%>
</td></tr>
		  
<tr class="forumRow">
<td align="right"><%=db_title%>地址：</td>
<td><input name="toUrl" type="text" id="toUrl" value="<%=toUrl%>" size="45" maxlength="255" /> 
(<span class="red">链接地址如: http://www.baidu.com</span>) </td>
</tr>
		  
<tr class="forumRow">
<td align="right" valign="top" style="padding-top:4px;">简要：</td>
<td><%=articlesHtml.noteBox("note",note)%></td>
</tr>
		  
<tr style="display:none" class="forumRaw">
<td align="right" valign="top">排序：</td>
<td><input name="order_id" type="text" id="order_id" value="<%=order_id%>" size="10" /></td>
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