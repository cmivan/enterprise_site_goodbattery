<!--配置文件-->
<!--#include file="article_config.asp"-->
<body>
<script language=javascript>
function ConfirmDel()
{
   if(confirm("确定要删除此用户吗？"))
     return true;
   else
     return false;
}
</script>
<TABLE border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<TR><td>

<!--#include file="../../models_system/articles/articles_manage_head.asp"-->


<%
dim id,Action,sql
    id=trim(Request("id"))
    Action=trim(request("Action"))
if id<>"" and isnumeric(id) then
	if Action="Lock" then
		sql="Update ["&db_table&"] set LockUser=true where id=" & CLng(id)
	elseif Action="CancelLock" then
		sql="Update ["&db_table&"] set LockUser=false where id=" & CLng(id)
    elseif Action="Del" then
	    sql="delete from ["&db_table&"] where id=" & Clng(id)
	end if
	if sql<>"" then conn.Execute sql  
end if


	'########## 处理接收到搜索字符的情况 #############
       keyword=request("keyword")
	if keyword<>"" then
	   keyword_sql =" and UserName like '%"&request("keyword")&"%'"
	   keyword_sql2=" where UserName like '%"&request("keyword")&"%'"
	end if
	
	'########## 定义排序字符 #############
	dim order_sql
	    order_sql=" order by id desc"

   '### 用于热门，审核 按钮链接,批量移动修改等
	Dim FullUrl
  	    FullUrl="?page="&page&"&keyword="&keyword&""

set rs=server.createobject("adodb.recordset")
    exec="select * from "&db_table&keyword_sql2&order_sql
	rs.open exec,conn,1,1 
	if rs.eof then
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="forum2 forumtop">
<tr class="forumRow">
<td align="center">
暂无 <%=db_title%> 相应内容!
</td>
</tr>
</table>	
	
<%else%>
<form name="FormE" method="post" action="" onSubmit="return CheckForm(this);">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="forum2 forumtop">
<tr class="forumRaw">
<td align="center" class="forumRaw"><strong> 序号</strong></td>
                  <td align="center" class="forumRaw"><strong> 用户名</strong></td>
                  <td align="center" class="forumRaw"><strong> 性别</strong></td>
                  <td align="center" class="forumRaw"><strong>E-mail</strong></td>
                  <td width="120" align="center" class="forumRaw"><strong>公司名称</strong></td>
                  <td align="center" class="forumRaw"><strong> 状态</strong></td>
                  <td width="120" align="center" class="forumRaw"><strong> 操作</strong></td>
</tr>
<%
	rs.PageSize =18 '每页记录条数
	iCount=rs.RecordCount '记录总数
	iPageSize=rs.PageSize
	maxpage=rs.PageCount 
	page=request("page")
	
	if Not IsNumeric(page) or page="" then
	   page=1
	else
	   page=cint(page)
	end if
	
	if page<1 then
	   page=1
	elseif  page>maxpage then
	   page=maxpage
	end if
	
	rs.AbsolutePage=Page
	if page=maxpage then
	   x=iCount-(maxpage-1)*iPageSize
	else
	   x=iPageSize
	end if	
	
For i=1 To x
%>
					
<tr class="forumRow">
<td align="center" bgcolor="#ECF5FF" class="forumRow"><%=rs("id")%></td>
                  <td align="center" bgcolor="#ECF5FF" class="forumRow"><%=rs("username")%></td>
                  <td align="center" bgcolor="#ECF5FF" class="forumRow"> 
                    <%if rs("Sex")=1 then
	    response.write "男"
	  else
	    response.write "女"
	  end if%>                  </td>
                  <td align="center" bgcolor="#ECF5FF" class="forumRow"><a href="../Email/EMAIL_Send.asp?email=<%=rs("email")%>"><%=rs("email")%></a></td>
                  <td width="120" align="center" bgcolor="#ECF5FF" class="forumRow"> <%=rs("Comane")%></td>
                  <td align="center" bgcolor="#ECF5FF" class="forumRow"> 
<%
	  if rs("LockUser")=true then
	  	response.write "<strong style=""color:#FF0000"">×</strong>"
	  else
	  	response.write "<strong style=""color:#00CC00"">√</strong>"
	  end if
	  %>                  </td>
                  <td width="120" align="center" bgcolor="#ECF5FF" class="forumRow"><a href="?Action=Del&id=<%=rs("id")%>&page=<%=page%>" onClick="return ConfirmDel();">删除</a>
                    <%if rs("LockUser")=False then %>
                    <a href="?Action=Lock&id=<%=rs("id")%>&page=<%=page%>">锁定</a> 
                    <%else%>
                    <a href="?Action=CancelLock&id=<%=rs("id")%>&page=<%=page%>">解锁</a> 
                    <%end if%>
                    <a href="article_edit.asp?id=<%=rs("id")%>&page=<%=page%>">修改</a></td>
					</tr>
				<%
					rs.movenext
					next
				%>
		</table>
			
</form>
				
		<%
			end if
		%>

<!--#include file="../../models_system/articles/articles_paging.asp"-->

</td></TR></TABLE>

<script>
function editItem(num){
 for (i=1;i<num;i++){
    Eitem=document.getElementById("edit_item_"+i).checked;
    if(Eitem==''){
    document.getElementById("edit_item_"+i).checked="checked";
	}else{
    document.getElementById("edit_item_"+i).checked='';
	}
 }
}
</script>



<%
if errStr="ok" then
   if edits="del" then
      errStr="成功删除!"
   elseif edits="check" then
      errStr="成功审核!"
   elseif edits="not_check" then
      errStr="成功取消审核!"
   elseif edits="move" then
      errStr="成功移动!"
   end if
end if

if errStr<>"" then
   response.Write("<script>alert('"&errStr&"');window.location.href='"&FullUrl&"';</script>")
end if
%>

</body>
</html>
<%
if request("act")="del" and db_showDel<>1 then
   conn.execute("delete * from "&db_table&" where id="&id)
   response.Write "<script>alert('成功刪除"&db_title&"！');window.location.href='"&FullUrl&"';</script>" 
end if
%>