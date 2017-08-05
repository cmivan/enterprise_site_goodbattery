<!--#include file="article_config.asp"-->
<!--#include file="../../libraries/articles.asp"-->
<%
dim typeB_id,typeS_id,keywords

	typeB_id = request.QueryString("typeB_id")
	typeS_id = request.QueryString("typeS_id")
	keywords = text.noSql(request("keywords"))

'//创建文章对象，并绑定数据库
set articles = New articlesClass
    articles.table = db_table
    articles.title = db_title
    '//设置文章是否被推荐等
    articles.setCheck()
    '//删除文章(单项)
    articles.del()
    '//批量操作(删除、移动、审核等)
    articles.batchEdit()
%>

<body>
<table border="0" align="center" cellpadding="0" cellspacing="10" class="forum1"><tr><td>

<!--#include file="../../models_system/articles/articles_manage_head.asp"-->

<%
set cmrs = New data
	if text.isnum(typeB_id) then cmrs.where "typeB_id",typeB_id
	if text.isnum(typeS_id) then cmrs.where "typeS_id",typeS_id
	if keywords<>"" then cmrs.likes "title",keywords
	cmrs.orderBy "order_id","asc"
	cmrs.orderBy "id","desc"
	cmrs.open db_table,12,true
	if cmrs.eof then
	   articles.noContent() '没有文章内容
	else
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop table">
<form name="news_category" method="post" class="form-search">

<tr class="forumRaw">
<td class="manage-ids">ID</td>
<td>&nbsp;&nbsp;<%=db_title%>&nbsp;名称\标题</td>
<td class="manage-edits">修改操作</td>
</tr>

<%do while not cmrs.eof%>

<tr class="forumRow">
<td class="manage-ids"><%=cmrs.rs("id")%></td>
<td>&nbsp;&nbsp;<a href="article_edit.asp?id=<%=cmrs.rs("id")%>"><%=text.keyColor( cmrs.rs("title") ,keywords)%></a></td>

<td class="manage-edits" id="<%=cmrs.rs("id")%>">
  <div class="btn-group"><a class="btn btn-mini delete" href="javascript:if(confirm('确定要删除此<%=db_title%>吗？删除后不可恢复!')){window.location.href='?act=del&id=<%=cmrs.rs("id")%>';}else{return false;}"><span class="icon icon-remove">&nbsp;</span> 删除</a><a class="btn btn-success btn-mini update" href="article_edit.asp?id=<%=cmrs.rs("id")%>"><span class="icon icon-white icon-edit">&nbsp;</span> 修改</a></div>
</td></tr>

<%cmrs.nexts:loop:cmrs.close%>			
</form>
</table>

<%
'page
'maxPage
'allCount
dim page,pageCounts,allCounts
page       = cmrs.page
pageCounts = cmrs.pageCounts
allCounts  = cmrs.allCounts
%>
<!--#include file="../../models_system/articles/articles_paging.asp"-->
				
<%end if%>

</td></tr></table>
</body></html>