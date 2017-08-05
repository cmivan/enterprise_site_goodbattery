<!--#include file="vote_config.asp"-->
<!--#include file="VOTE_TOP.ASP"-->
<%set rs=server.CreateObject("adodb.recordset")
 sql="delete * from vote_title where id="&request("id")
 rs.open sql,connstr
 set rs=server.CreateObject("adodb.recordset")
 sql="delete * from vote_choice where extends="&request("id")
 rs.open sql,connstr
 response.write "<script language='javascript'>alert('删除成功！');window.location.href='VOTE_ADMIN.ASP';</script>"
%>