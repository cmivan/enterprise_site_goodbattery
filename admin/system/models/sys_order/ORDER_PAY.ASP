<!--#include file="../SYS_INCLUDE/SYS_INCLUDE.ASP"-->



<%
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = Conn
Rs.Open "SELECT * FROM ORDER_PayType Order BY id"

if Request.QueryString("mark")="yes" then
id= Trim(Request("id"))
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = Conn
Rs.Open "DELETE FROM ORDER_PayType Where id="&id,Conn,2,3,1
Set Rs= Nothing
Set Conn = Nothing
Response.Redirect "ORDER_PAY.ASP"
end if

if Request.QueryString("mark")="southidc" then
paytype=request.form("paytype")
paymentmessage=request.form("paymentmessage")

If paytype="" Then
Response.Write("<script language=""JavaScript"">alert(""错误：您没输入支付银行，请返回检查！！"");history.go(-1);</script>")
response.end
end if

If paymentmessage=""Then
Response.Write("<script language=""JavaScript"">alert(""错误：您没有输入支付说明，请返回检查！！"");history.go(-1);</script>")
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from ORDER_PayType "
rs.open sql,conn,1,3
rs.addnew
rs("paytype")=paytype
rs("paymentmessage")=paymentmessage
rs.update
Response.Redirect "ORDER_PAY.ASP"
end if
%>
<br />

<table width="600" border="0" align=center cellpadding=10 cellspacing=1 bordercolor="#FFFFFF" bgcolor="#C4D8ED">
        <tr> 
          <form method="post" action="ORDER_PAY.ASP?mark=southidc">
            <td bgcolor="F2F8FF"> 
<table width="100%" border="0" align=center cellpadding=2 cellspacing=1 bordercolor="#FFFFFF" bgcolor="#C4D8ED">
                <tr> 
                  <td height="24" colspan="2" align="center" class="forumRaw"><strong>支付管理</strong></td>
                </tr>
                <tr> 
                  <td width="125" height="24" align="center" class="forumRow"> 
                    支付银行</td>
                  <td bgcolor="#ECF5FF" class="forumRow">
<input name="paytype" type="text" id="paytype" size="40" maxlength="40"></td>
                </tr>
				
                <tr> 
                  <td width="125" height="22" align="center" class="forumRow">
支付说明</td>
                  <td bgcolor="#ECF5FF" class="forumRow">
<textarea name="paymentmessage" cols="60" rows="5" id="paymentmessage"></textarea></td>
                </tr>
                <tr> 
                  <td height="22" class="forumRow">&nbsp;</td>
                  <td height="22" class="forumRow"><input type="submit" name="Submit" value="提交保存" /></td>
                </tr>
              </table>
              <br> 
<table width="100%" border="0" align=center cellpadding=3 cellspacing=1 bordercolor="#FFFFFF" bgcolor="#C4D8ED">
                <tr> 
                  <td width="125" height="25" align="left" class="forumRaw"> 
                  &nbsp;支付银行</td>
                  <td align="left" class="forumRaw"> 
                  &nbsp;支付说明</td>
                  <td colspan="2" align="center" class="forumRaw"> 操作</td>
                </tr>
                <% do while not Rs.eof %>
                <tr bgcolor="#DFEBF2"> 
                  <td width="125" height="22" align="left" class="forumRow"><%=rs("paytype")%></td>
                  <td align="left" class="forumRow"><%=Rs("paymentmessage")%></td>
                  <td width="35" align="center" class="forumRow"> 
                  <a href="ORDER_EDITPAY.ASP?id=<%=Rs("id")%>">修改</a></td>
                  <td width="35" align="center" class="forumRow"> 
                  <a href="ORDER_PAY.ASP?id=<%=Rs("id")%>&amp;mark=yes">删除</a></td>
                </tr>
                <%Rs.MoveNext 
loop 
%>
              </table>
              <%  Set Rs = Nothing 
Set Conn = Nothing 
%> </td>
          </form>
        </tr>
      </table>
<br />
