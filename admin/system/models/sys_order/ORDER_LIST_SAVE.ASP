<%@ language="VBScript"%>
<!--#include file="../SYS_INCLUDE/SYS_INCLUDE.ASP"-->
<%
djfc="已经发货"
set rs=server.createobject("adodb.recordset")
sqltext="select Flag from ORDER_List where OrderNum='"&request("OrderNum")&"'"
rs.open sqltext,conn,1,1
if rs("Flag")="已经发货" then
rs.close
Response.Redirect "ORDER_LIST_SAVE.ASP?msg=此订货单您已经做过了发货处理！"
Else
set rs=server.createobject("adodb.recordset")
sqltext="update ORDER_List set Flag='"&djfc&"' where OrderNum='"&request("OrderNum")&"'"
rs.open sqltext,conn,3,3
response.redirect "ORDER_MESSAGE_BOX.ASP?msg=订货单处理处理完毕，请按客户详细地址发货！"
end if
%>