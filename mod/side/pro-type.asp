<%if is_side_proType=true then%>

<div class="mod">
<div class="m-body">
<div class="m-header"><h3>产品分类</h3></div>
<div class="m-content">
      
<%
set t1Rs = New data
    t1Rs.where "type_id",0
	t1Rs.orderBy "order_id","asc"
	t1Rs.orderBy "id","desc"
	t1Rs.open "product_type","*",false
	if not t1Rs.eof then
%>
<div class="shop-category">
<div class="bd">
<ul class="vas">

<%
do while not t1Rs.eof

	'//获取子分类
	isT2Rs = false
	set t2Rs = New data
	    t2Rs.where "type_id",t1Rs.rs("id")
		t2Rs.orderBy "order_id","asc"
		t2Rs.orderBy "id","desc"
		t2Rs.open "product_type","*",false
		if not t2Rs.eof then
		   isT2Rs = true
		end if
%>

<%if isT2Rs then%>
  <li class="cat expand">
  <h4 class="cat-hd"><i></i><a href="offerlist.asp<%=URI.reUrl("typeB_id="&t1Rs.rs("id"))%>" title="<%=t1Rs.rs("title")%>"><span><%=text.strCut(t1Rs.rs("title"),20)%></span></a></h4>
  <ul class="cat-bd">
    <%do while not t2Rs.eof%>
    <li><a href="offerlist.asp<%=URI.reUrl("typeB_id="&t1Rs.rs("id")&"&typeS_id="&t2Rs.rs("id"))%>" title="<%=t2Rs.rs("title")%>"><span><%=t2Rs.rs("title")%></span></a></li>
    <%t2Rs.nexts:loop:t2Rs.close%>
  </ul>
  </li>
<%else%>
  <li class="cat last"><h4 class="cat-hd"><i></i><a href="offerlist.asp<%=URI.reUrl("typeB_id="&t1Rs.rs("id"))%>" title="<%=t1Rs.rs("title")%>"><span><%=text.strCut(t1Rs.rs("title"),20)%></span></a></h4></li>
<%end if%>

<%
	set t2Rs = nothing
t1Rs.nexts:loop:t1Rs.close
%>

</ul>
</div>
</div>
<%
	end if
set t1Rs = nothing
%>

</div></div></div>
<%end if%>