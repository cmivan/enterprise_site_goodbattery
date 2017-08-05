<%DB_PATH="../"%>
<!--#include file="../Inc/Conn.asp"-->

<%
if isempty(session("go")) then  '使用session记录是否已投票， 使用时容易被刷
   choice=request("choice")
   if isempty(choice) then
response.write"<script language='javascript'>alert('您还没选择投票内容！');history.go(-1);</script>"
   else
      set rs=server.CreateObject("adodb.recordset")
      sql="update vote_choice set num=num+1 where id in("&choice&")"
      rs.open sql,connstr
      session("go")=1
      response.write "<script language='javascript'>alert('投票成功');</script>"
      response.write "<meta http-equiv='refresh' content='1;url=Show.asp?id="&request("id")&"'>"
   end if
else
   response.write "<script language='javascript'>alert('您已经投过票了！');history.go(-1)</script>"
end if
   
%>