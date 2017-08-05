<%@ LANGUAGE="VBScript" %>
<!--#include file="../SYS_INCLUDE/SYS_INCLUDE.ASP"-->
<%
dim SQL, Rs, contentID,CurrentPage
CurrentPage = request("Page")
contentID=clng(request("ID"))

set rs=server.createobject("adodb.recordset")
sqltext="delete from Order_List where id="&contentID
rs.open sqltext,conn,3,3
set rs=nothing

set rs=server.createobject("adodb.recordset")
sqltext="delete from ShopList where id="&contentID
rs.open sqltext,conn,3,3
set rs=nothing
conn.close

response.redirect "ORDER_LIST.ASP?page="&CurrentPage
%>