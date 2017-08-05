<!--#include file="../SYS_INCLUDE/SYS_INCLUDE.ASP"-->


<%if Request.QueryString("no")="eshop" then

id=request("id")
paytype=request("paytype")
paymentmessage=Request("paymentmessage")

If paytype="" Then
response.write "SORRY <br>"
response.write "请输入支付银行"
response.end
end if
If paymentmessage="" Then
response.write "SORRY <br>"
response.write "支付说明不能为空"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from ORDER_PayType where id="&id
rs.open sql,conn,1,3

rs("paytype")=paytype
rs("paymentmessage")=paymentmessage
rs.update
rs.close
response.redirect "ORDER_PAY.ASP"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From ORDER_PayType where id="&id, conn,3,3
%>

<br />
<table width="600" border="0" align=center cellpadding=10 cellspacing=1 bordercolor="#FFFFFF" bgcolor="#C4D8ED">
        <tr> 
          <form method="post" action="ORDER_EDITPAY.ASP?no=eshop">
            <input type=hidden name=id value=<%=id%>>
            <td class="forumRow">
<table width="100%" border="0" align=center cellpadding=2 cellspacing=1 bordercolor="#FFFFFF" bgcolor="#C4D8ED">
              <tr> 
                    <td colspan="2" bgcolor="#A6EDFF" class="forumRaw"><div align="center"><strong>修改支付</strong></div>                      </td>
          </tr>
				                    <tr> 
                    <td width="125" height="25" valign="top" bgcolor="#A6EDFF" class="forumRow"> 
                      <div align="center">支付银行</div></td>
                    <td bgcolor="#E3E3E3" class="forumRow">
<input name="paytype" type="text" id="paytype" value="<%=rs("paytype")%>" size="40" maxlength="40"></td>
                  </tr>
                  <tr> 
                    <td width="125" height="22" valign="top" bgcolor="#A6EDFF" class="forumRow">
<div align="center">支付说明</div></td>
                    <td bgcolor="#E3E3E3" class="forumRow"><textarea name="paymentmessage" cols="60" rows="5" id="paymentmessage"><%=rs("paymentmessage")%></textarea></td>
                  </tr>
                  <tr bgcolor="#A6EDFF"> 
                    <td height="25" class="forumRaw">&nbsp;                      </td>
                    <td height="25" class="forumRaw"><input name="submit" onclick="history.back(1);" type="button" value="返回" />
&nbsp;
<input name="submit2" type="submit" value="提交保存" /></td>
                  </tr>
              </table>
            </td>
          </form>
        </tr>
      </table>
<br />

