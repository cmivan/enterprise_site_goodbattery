<!--#include file="vote_config.asp"-->
<!--#include file="VOTE_TOP.ASP"-->
<%
cid=request("cid")
set rs=server.CreateObject("adodb.recordset")
sql="delete * from vote_choice where id="&cid
rs.open sql,connstr
 response.write "<script language='javascript'>alert('删除成功！');window.location.href='VOTE_ADD_CHOICE.ASP?id="&request("id")&"';</script>"
%>
