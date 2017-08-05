<%
function product_list(num)

	dim typeB_id,typeS_id,keyword
	
		typeB_id = request.QueryString("typeB_id")
		typeS_id = request.QueryString("typeS_id")
		keyword = text.noSql(request("keyword"))
		
	set myrs = New data
	    'myrs.where "ok",1
		
		if keyword<>"" then myrs.likes "title",keyword
		if text.isnum(typeS_id) then myrs.where "typeS_id",typeS_id
		if text.isnum(typeB_id) then myrs.where "typeB_id",typeB_id

		myrs.orderBy "order_id","asc"
		myrs.orderBy "id","desc"
		
		myrs.open "product",num,true

	set product_list = myrs
	
end function
%>