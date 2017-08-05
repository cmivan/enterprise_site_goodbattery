<!--#include file="vote_config.asp"-->
<body>
<table border="0" align=center cellpadding=0 cellspacing=10 class="forum1">
<tr><td>
<!--#include file="VOTE_TOP.ASP"-->

<%set rs=server.CreateObject("adodb.recordset")
  sql="select * from vote_title order by id desc"
  rs.open sql,connstr
%>

<%if rs.eof and rs.bof then%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forum2 forumtop">
<tr class="forumRaw"><td width="44%" height="21">
目前还没有投票</td></tr></table>
<%else%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forum2">
<form name="form1" method="post" action="">

<%
do while not rs.eof
       if rs("current")=true then
	      current="激活"
	   else
	      current="未激活"
	   end if
  %>
<tr class="forumRaw">
      <td width="44%" height="21">&nbsp;<a href="VOTE_EDIT.ASP?id=<%=rs("id")%>"><%=rs("title")%></a></td>
      <td width="12%"><div align="center"><a href="VOTE_ADD_CHOICE.ASP?id=<%=rs("id")%>">添加选项</a></div></td>
      <td width="7%"><div align="center"><a href="VOTE_DEL.ASP?id=<%=rs("id")%>" onClick="return test()">删除</a>
      </div></td>
      <td width="9%"><div align="center">
	  <a href="../../VOTE/VOTE.ASP?id=<%=rs("id")%>" target="_blank">预览</a></div></td>
      <td width="28%"><div align="center">
        <input name="textfield" type="text" class="input" value="&lt;object type=&quot;text/x-scriptlet&quot; width=&quot;400&quot; height=&quot;300&quot; data=&quot;vote.asp?id=<%=rs("id")%>&quot;&gt;&lt;/object&gt;" size="22">
      </div></td>
    </tr>
	<%rs.movenext
	  loop
	  rs.close
	  set rs=nothing
	  %>
      
</form>
</table>

<%end if%>

</td></tr></table>

</body>

