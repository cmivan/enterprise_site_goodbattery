<!--#include file="vote_config.asp"-->
<body>
<table border="0" align=center cellpadding=0 cellspacing=10 class="forum1">
<tr><td>
<!--#include file="VOTE_TOP.ASP"-->

<%if request("actions")="1" then
      choice=trim(request("choice"))
 if choice="" then
	      response.write "<script language='javascript'>alert('选项内容不能为空!');history.go(-1);</script>"
      else
	     if request("IsDefault")="1" then
		    def=true
         else
		    def=false
		 end if 
	     set rs=server.CreateObject("adodb.recordset")
         sql="select * from vote_choice where id=null"
         rs.open sql,connstr,1,3
		 rs.addnew
         rs("choice")=choice
		 rs("IsDefault")=def
		 rs("extends")=request("ids")
		 rs.update
		 rs.close
		 set rs=nothing
		 response.write "<meta http-equiv='refresh' content='0;url=VOTE_ADD_CHOICE.ASP?id="&request("ids")&"'>"
      end if
else
set rs=server.CreateObject("adodb.recordset")
sql="select * from vote_title where id="&request("id")
rs.open sql,connstr
if rs("choice")=true then
   con="多选"
else
   con="单选"
end if
if rs("current")=true then
   curr="激活"
else
   curr="未激活"
end if
%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forum2 forumtop">
<form name="form1" method="post" action="VOTE_EDIT.ASP">
<tr class="forumRaw">
      <td height="20" colspan="2">&nbsp;当前主题内容</td>
    </tr>
<tr class="forumRow">
      <td width="442" valign="middle">&nbsp;<%=rs("title")%></td>
      <td width="147" align="center" valign="middle">类型：<%=con%>&nbsp; <a href="VOTE_EDIT.ASP?id=<%=rs("id")%>">修改</a>
        <input name="id" type="hidden" id="id" value="<%=rs("id")%>" /></td>
</tr>
</form>
</table>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forum2">
<form name="form2" method="post" action="">
<tr class="forumRaw">
      <td>&nbsp;选项内容</td>
      <td align="center"> &nbsp;是否默认选中</td>
      <td align="center">编辑</td>
    </tr>
<tr class="forumRow">
      <td width="65%" valign="middle" class="forumRaw"><input name="choice" type="text" class="input" id="choice" size="30"></td>
      <td width="16%" align="center" valign="middle" class="forumRaw">
        <input type="radio" name="IsDefault" value="1">
        是
        <input name="IsDefault" type="radio" value="2" checked>
      否</td>
      <td width="19%" valign="middle" class="forumRaw"><div align="center">
        <input name="ids" type="hidden" id="ids" value="<%=request("id")%>">
        <input name="actions" type="hidden" id="actions" value="1">
        <input type="submit" name="Submit" value="提交保存" class="button" />
      </div></td>
    </tr>
</form>


<%set rs1=server.CreateObject("adodb.recordset")
sql1="select * from vote_choice where extends="&request("id")
rs1.open sql1,connstr%>
 <%if rs1.eof and rs1.bof then
   else
   do while not rs1.eof
   if rs1("IsDefault")=true then
     Def="是"
   else
     Def="否"
   end if%>
<tr class="forumRow">
    <td height="21" valign="middle">&nbsp;<%=rs1("choice")%></td>
    <td align="center" valign="middle"><%=Def%></td>
    <td valign="middle" ><div align="center"><a href="VOTE_EDIT_CHOICE.ASP?cid=<%=rs1("id")%>&id=<%=request("id")%>">修改</a> <a href="VOTE_DEL_CHOICE.ASP?cid=<%=rs1("id")%>&id=<%=request("id")%>" onClick="return test()">删除</a></div></td>
  </tr>

  <%rs1.movenext
    loop
	rs1.close
	set rs1=nothing
	end if%>
<tr class="forumRaw">
   <td height="21" colspan="3" valign="middle" ><div align="center">
     <%rs.close
set rs=nothing
end if%>
     <font color="red">注意：请确保每个主题只有一个默认选项，否则将出现不可预料的结果！</font> </div></td>
   </tr>
</table>


	</td>
  </tr>
  
</table>

</body>