<!--#include file="article_config.asp"-->
<%

'///////////  处理提交数据部分 ////////// 
   edit_id   =request("id")
if edit_id="" or isnumeric(edit_id)=false then
   editStr="添加"
   else
   editStr="修改"
end if
if request.Form("edit")="ok" then
'///////////  写入数据部分 //////////
   set rs=server.createobject("adodb.recordset") 
    if edit_id="" or isnumeric(edit_id)=false then
       exec="select * from "&db_table                       '判断，添加数据
	   rs.open exec,conn,1,3
       rs.addnew
	else
	   exec="select * from "&db_table&" where id="&edit_id  '判断，修改数据
       rs.open exec,conn,1,3
	end if
	if edit_id<>"" and isnumeric(edit_id) and rs.eof then
	   response.Write("写入数据失败!")
	else

		dim UserName,Password,Question,Answer,Sex,Email,Homepage,LockUser,Comane,Add,Somane,Zip,Phone,Fox
			UserName=trim(request("UserName"))
			Password=trim(request("Password"))
			Question=trim(request("Question"))
			Answer=trim(request("Answer"))
			Sex=trim(Request("Sex"))
			Email=trim(request("Email"))
			Homepage=trim(request("Homepage"))
			Comane=trim(request("Comane"))
			Add=trim(request("Add"))
			Somane=trim(request("Somane"))
			Zip=trim(request("Zip"))
			Phone=trim(request("Phone"))
			Fox=trim(request("Fox"))
			LockUser=request("LockUser")

if UserName="" or len(UserName)>14 or len(UserName)<4 then
				founderr=true
				errmsg=errmsg & "请输入用户名(不能大于14小于4)"
			else
  				if Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"")>0 or Instr(UserName,"$")>0 then
					errmsg=errmsg+"用户名中含有非法字符"
					founderr=true
				else
				dim sqlReg,rsReg
                    if edit_id<>"" and isnumeric(edit_id) then
					sqlReg="select * from [User] where UserName='" & Username & "' and id<>" & edit_id
                    else
					sqlReg="select * from [User] where UserName='" & Username & "'"
                    end if
					set rsReg=server.createobject("adodb.recordset")
					rsReg.open sqlReg,conn,1,1
					if not(rsReg.bof and rsReg.eof) then
						founderr=true
						errmsg=errmsg & "用户名已经存在！请换一个用户名再试试！"
					end if
					rsReg.Close
					set rsReg=nothing
				end if
			end if
			if Password<>"" then
				if len(Password)>12 or len(Password)<6 then
					founderr=true
					errmsg=errmsg & "请输入密码(不能大于12小于6)。如不想修改，请留空！"
				else
					if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"")>0 or Instr(Password,"$")>0 then
						errmsg=errmsg+"密码中含有非法字符"
						founderr=true
					end if
				end if
			end if
			if Question="" then
				founderr=true
				errmsg=errmsg & "密码提示问题不能为空"
			end if
			if Sex="" then
				founderr=true
				errmsg=errmsg & "性别不能为空"
			else
				sex=cint(sex)
				if Sex<>0 and Sex<>1 then
					Sex=1
				end if
			end if
			if Email="" then
				founderr=true
				errmsg=errmsg & "Email不能为空"
			end if

			if LockUser="1" then
			   LockUser=True
            else
			   LockUser=False
			end if

			if FoundErr<>true then
				rs("UserName")=UserName
				if Password<>"" then
					rs("Password")=md5(Password)
				end if
				rs("Question")=Question
				if Answer<>"" then
					rs("Answer")=md5(Answer)
				end if
				rs("Sex")=Sex
				rs("Email")=Email
				rs("HomePage")=HomePage
				rs("Comane")=Comane
				rs("Add")=Add
				rs("Somane")=Somane
				rs("Zip")=Zip
				rs("Phone")=Phone
				rs("Fox")=Fox
				rs("LockUser")=LockUser

            if sys_back then sys_backs="window.location.href='article_manage.asp';"
               response.Write("<script>alert('"&editStr&"操作成功!');"&sys_backs&"</script>")
	     end if
	rs.update
	rs.close

			end if

set rs=nothing
end if

'#################  读取数据部分 #################
   id=request.QueryString("id") 
if id<>"" and isnumeric(id) then
   edit_stat="edit"
   set rs=server.createobject("adodb.recordset") 
       exec="select * from "&db_table&" where id="&edit_id
       rs.open exec,conn,1,1 
	   if not rs.eof then
			UserName=rs("UserName")
			Password=rs("Password")
			Question=rs("Question")
			Answer=rs("Answer")
			Sex=rs("Sex")
			Email=rs("Email")
			Homepage=rs("Homepage")
			Comane=rs("Comane")
			Add=rs("Add")
			Somane=rs("Somane")
			Zip=rs("Zip")
			Phone=rs("Phone")
			Fox=rs("Fox")
			LockUser=rs("LockUser")
	   end if
	   rs.close
   set rs=nothing
else
   edit_stat="add"  '/// 当前状态:添加
end if


  '/// 处理添加时间
  if isdate(add_data)=false then add_data=now()
%>


<body>
<TABLE border="0" align="center" cellpadding="0" cellspacing="10" class="forum1">
<TR><td>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="forum2 forumtop">
<tr class="forumRaw">
<td align="center"><strong>&nbsp;&nbsp;<%=db_title%> <%=editStr%></strong></td>
</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forum2 table">
<form name="article_update" method="post" action="">
<tr class="forumRow">
<TD width="120" align="right">用户名称：</TD>
<TD> <INPUT name=UserName value="<%=UserName%>" size=30   maxLength=14></TD>
</TR>

<tr class="forumRow">
<TD width="120" align="right">密码(至少6位)：</TD>
<TD><INPUT   type=password maxLength=16 size=30 name=Password> <font color="#FF0000">如果不想修改，请留空</font>            </TD>
</TR>

<tr class="forumRow"> 
<TD width="120" align="right">密码问题：</TD>
<TD> <INPUT name="Question"   type=text value="<%=Question%>" size=30>            </TD>
</TR>

<tr class="forumRow"> 
<TD width="120" align="right">问题答案：</TD>
<TD> <INPUT   type=text size=30 name="Answer"> <font color="#FF0000">如果不想修改，请留空</font></TD>
</TR>

<tr class="forumRow"> 
<TD width="120" align="right">性 别：</TD>
<TD> <INPUT type=radio value="1" name=sex <%if Sex=1 then response.write "CHECKED"%>>
              男 &nbsp;&nbsp; <INPUT type=radio value="0" name=sex <%if Sex=0 then response.write "CHECKED"%>>
            女</TD>
</TR>

<tr class="forumRow"> 
<TD width="120" align="right">E-mail：</TD>
<TD> <INPUT name=Email value="<%=Email%>" size=30   maxLength=50>            </TD>
</TR>

<tr class="forumRow"> 
<TD width="120" align="right">个人主页：</TD>
<TD> <INPUT   maxLength=100 size=30 name=homepage value="<%=HomePage%>"></TD>
</TR>

<tr class="forumRow"> 
<TD width="120" align="right">公司名称：</TD>
<TD> <INPUT name=Comane value="<%=Comane%>" size=30 maxLength=20></TD>
</TR>

<tr class="forumRow"> 
<TD width="120" align="right">收货地址：</TD>
<TD> <INPUT name=Add value="<%=Add%>" size=30 maxLength=50></TD>
</TR>

<tr class="forumRow"> 
<TD align="right">收货人：</TD>
<TD><input name=Somane value="<%=Somane%>" size=30 maxlength=50></TD>
</TR>

<tr class="forumRow"> 
<TD align="right">邮政编码：</TD>
<TD><INPUT name=Zip value="<%=Zip%>" size=30 maxLength=50></TD>
</TR>

<tr class="forumRow"> 
<TD align="right">联系电话：<br></TD>
<TD><INPUT name=Phone value="<%=Phone%>" size=30 maxLength=50></TD>
</TR>

<tr class="forumRow"> 
<TD align="right">传真号码：</TD>
<TD><INPUT name=Fox value="<%=Fox%>" size=30 maxLength=50></TD>
</TR>

<tr class="forumRow">
<TD width="120" align="right">用户状态：</TD>
<TD>
<%if LockUser=True then%>
<input name="LockUser" type="checkbox" value="1" checked />
(正常)&nbsp;&nbsp;
<%else%>
<input name="LockUser" type="checkbox" value="1" />
(锁定)
<%end if%></TD>
          </TR>
		  
<tr class="forumRaw">
<td align="right" valign="top">
<input name="id" type="hidden" value="<%=id%>" />
<input name="edit" type="hidden" value="ok" /></td>
<td>
<%=articlesHtml.submitBtn("",editStr)%>
<%
if errmsg<>"" then response.Write("<span class='red'>"&errmsg&"</span>")
%></td></tr>
</form>
</table>
</td></TR></TABLE>
</body>
</html>