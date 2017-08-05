<!--#include file="../SYS_INCLUDE/SYS_INCLUDE.ASP"-->

<%
id=request("id")
page=request("page")
set rs=server.createobject("adodb.recordset")
sqltext="select * from ORDER_List where OrderNum='"&id&"'"
rs.open sqltext,conn,1,1
%>
<br />
<table width="550" border="0" align=center cellpadding=10 cellspacing=1 bordercolor="#FFFFFF" bgcolor="#C4D8ED"><tr> <td class="forumRow">
<form method='POST' style="margin:0;" action="ORDER_LIST_SAVE.ASP?OrderNum=<%=rs("OrderNum")%>">
<table width="100%" border="0" align=center cellpadding=3 cellspacing=1 bordercolor="#FFFFFF" bgcolor="#C4D8ED">
                <TBODY>
                  <TR> 
                    <TD width="548"  colSpan=2 align="center" valign="top" bgcolor="#DBDBDB" class="forumRaw"><strong><%=rs("OrderNum")%>&nbsp;号订单管理 </strong>客户订货单详细资料</TD>
                  </TR>
             
                  <TR> 
                    <TD  width=126 height=25 align="right" bgColor=#DBDBDB class="forumRow"><font color="#000000">订货单号：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee" class="forumRow"> <%=rs("OrderNum")%>&nbsp;</TD>
                  </TR>
                  <tr> 
                    <TD  width=126 height=25 align="right" bgColor=#DBDBDB class="forumRow">公司名称<font color="#000000">：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee" class="forumRow"> <%=rs("Comane")%>&nbsp;</TD>
                  </TR>
                  <tr> 
                    <TD  width=126 height=25 align="right" bgColor=#DBDBDB class="forumRow"><font color="#000000">
					收 货 人：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee" class="forumRow"> <%=rs("RecName")%>&nbsp;</TD>
                  </tr>
                  <tr> 
                    <TD  width=126 height=25 align="right" bgColor=#DBDBDB class="forumRow"><font color="#000000">收货地址：</font></TD>
                    <TD  width=410 height=-2 bgcolor="#eeeeee" class="forumRow"> <%=rs("RecAddress")%>&nbsp;</TD>
                  </tr>
                  <tr> 
                    <TD height=25 align="right" bgColor=#DBDBDB class="forumRow"><font color="#000000">邮政编码：</font></TD>
                    <TD  width=410 height=0 bgcolor="#eeeeee" class="forumRow">  <%=rs("ZipCode")%>&nbsp;</TD>
                  </tr>
                  <tr> 
                    <TD height=25 align="right" bgColor=#DBDBDB class="forumRow"><font color="#000000">联系电话：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee" class="forumRow"> <%=rs("RecPhone")%>&nbsp;</TD>
                  </tr>
                  <tr> 
                    <TD  width=126 height=25 align="right" bgColor=#DBDBDB class="forumRow">联系传真<font color="#000000">：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee" class="forumRow"> <%=rs("Fox")%>&nbsp;</TD>
                  </tr>
                  <tr> 
                    <TD  width=126 height=25 align="right" bgColor=#DBDBDB class="forumRow"><font color="#000000">电子信箱：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee" class="forumRow"> <%=rs("RecMail")%>&nbsp;</TD>
                  </tr>
                  <tr> 
                    <TD  width=126 height=25 align="right" bgColor=#DBDBDB class="forumRow"><font color="#000000">所选
					付款帐号：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee" class="forumRow"> <%=rs("PayType")%>&nbsp;</TD>
                  </tr>
                  <tr> 
                    <TD  width=126 height=25 align="right" bgColor=#DBDBDB class="forumRow"><font color="#000000">备 注：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee" class="forumRow"> <%=rs("Notes")%>&nbsp;</TD>
                  </tr>
                  <tr> 
                    <TD  width=126 height=24 align="right" bgColor=#DBDBDB class="forumRow"><font color="#000000">订货日期：</font></TD>
                    <TD  width=410 height=24 bgcolor="#eeeeee" class="forumRow"> <%=rs("OrderTime")%>&nbsp;</TD>
                  </tr>
                  <tr> 
                    <TD  width=126 height=25 align="right" bgColor=#DBDBDB class="forumRow"><font color="#000000">订单是否已经处理：</font></TD>
                    <TD  width=410 height=25 bgcolor="#eeeeee" class="forumRow"> <%If rs("Flag")="尚未处理" Then%>
                      尚未处理 
                        <%else%>
                      已经发货 
                    <%End If%>&nbsp; </TD>
                  </tr>
                  <TR> 
                    <TD width="548"  colSpan=2 valign="top" class="forumRaw">订货商品细目</TD>
                  </TR>
<%
set rs2=server.createobject("adodb.recordset")

sqltext2 = "select A.OrderNum,A.Product_Id,A.BuyPrice,A.ProdUnit,"&_		
		"C.Title,C.Buy_Price "&_
		" from ORDER_Detail A,Product_ C"&_
		" where A.OrderNum='"&id&"' and A.Product_Id=C.Product_Id"
		
'sqltext2="select * from ORDER_Detail where OrderNum='"&id&"'"
rs2.open sqltext2,conn,1,1
%>&nbsp;
                  <TR> 
                    <TD width="548" height="15"  colSpan=2 valign="top" bgcolor="#eeeeee" class="forumRow"> 
                      <div align="center"> 
                        <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#006699" bordercolordark="#eeeeee"  height="67">
                          <tr> 
                            <td width="49%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">产品名称</font></td>
                            <td width="14%" bgcolor="#DBDBDB" height="21" align="center"><font color="#000000">购买数量</font></td>
                            <td width="18%" bgcolor="#DBDBDB" height="21" align="center">购买单价</td>
                            <td width="19%" height="21" align="center" bgcolor="#DBDBDB">合 
                              计</td>
                          </tr>
                          <% Total = 0
						  BuyTotal=0
While Not rs2.EOF
     Sum=0
	 BuySum=0
     Sum = rs2("ProdUnit")*  FormatNumber(rs2("BuyPrice"),1)
	 BuySum = rs2("ProdUnit")*  FormatNumber(rs2("Buy_Price"),1)
	 Sum = FormatNumber(Sum,1) 
	 BuySum = FormatNumber(BuySum,1) 
	 Total = Sum + Total 
	 BuyTotal = BuySum + BuyTotal '计算总金额%>&nbsp;
                          <tr> 
                            <td width="49%" align="center" height="22"><%=rs2("Title")%>&nbsp;</td>
                            <td width="14%" align="center" height="22"><%=rs2("ProdUnit")%>&nbsp;</td>
                            <td width="18%" align="center" height="22"><%=rs2("BuyPrice")%>&nbsp; </td>
                            <td height="22" align="center"><%=Sum%>&nbsp;</td>
                          </tr>
                          <%
rs2.MoveNext
Wend
%>&nbsp;
                          <tr> 
                            <td colspan="2" height="22"> <p align="left">进价总金额:<%=FormatNumber(BuyTotal,2)%>&nbsp;</p></td>
                            <td height="22" colspan="2">&nbsp;售价总金额:<%=FormatNumber(Total,2)%>&nbsp;</td>
                          </tr>
                        </table>
                    </div></TD>
                  </TR>
                <center>
                  <TR> 
                    <TD width="548" height="25"  colSpan=2 class="forumRaw"> 
                      <p align="center"> 
                        <%If rs("Flag")="尚未处理" Then%>&nbsp;
                        <input type="submit" value="订单处理" name="B1">
                        <%
rs.close
rs2.close
conn.close
End If
%>&nbsp;
<input type="button" value="返回" name="B4" onclick="javascript:window.history.go(-1)">
                    </TD>
                  </TR>
                  <TR bgColor=#DBDBDB> 
                    <TD width="548" height="3"  colSpan=2></TD>
                  </TR>
                </center>
            </TABLE>
 </form>
</td>
         
        </tr>
</table>
<br />
