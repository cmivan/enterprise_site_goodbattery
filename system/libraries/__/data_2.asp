<%
'--操作数据库类data
Class data

 '--数据集操作对象
 public  rs
 Private rsSelect
 Private rsFrom
 Private rsQuery
 Private rsWhere
 Private rsOrder
 public rsNum
 '----------
 public page
 public pageSizes
 public pageCounts
 public allCounts

 Private Sub Class_Initialize  '--类初始化
  Set rs = Server.CreateObject("Adodb.Recordset")
      rsSelect = "*"
	  rsFrom   = ""
      rsQuery  = ""
	  rsWhere  = ""
	  rsOrder  = ""
      rsNum  = 0
	  page = 1
      pageSizes = 10
	  pageCounts = 0
	  allCounts = 0
 End Sub
 
 'Select /////////////
 public Sub selects(key,num)
      if key<>"" then
	     rsSelect = key
		 if isnumeric(num) and num>0 then rsSelect = "top " & num & " " & rsSelect
	  end if
 End Sub
 
 'From /////////////
 public Sub from(key)
      if key<>"" then rsFrom = key
 End Sub
 
 'Where /////////////
 public function whereTemp(key,val)
      if key<>"" then whereTemp = "(" & key & " = " & val & ")"
 End function
 
 public Sub where(key,val)
      dim whereTemps
      if key<>"" then
	     whereTemps = whereTemp(key,val)
	     if rsWhere<>"" then
		    rsWhere = rsWhere & " and " & whereTemps
		 else
		    rsWhere = whereTemps
		 end if
	  end if
 End Sub
 
 public Sub whereOr(key,val)
      dim whereTemps
      if key<>"" then
	     whereTemps = whereTemp(key,val)
	     if rsWhere<>"" then
		    rsWhere = rsWhere & " or " & whereTemps
		 else
		    rsWhere = whereTemps
		 end if
	  end if
 End Sub

 'Like /////////////
 public function likesTemp(key,val)
      if key<>"" then likesTemp = "(" & key & " like '%" & val & "%')"
 End function
 
 public Sub likes(key,val)
      dim likeTemps
      if key<>"" and val<>"" then
	     likeTemps = likesTemp(key,val)
	     if rsWhere<>"" then
		    rsWhere = rsWhere & " and " & likeTemps
		 else
		    rsWhere = likeTemps
		 end if
	  end if
 End Sub
 
 public Sub likesOr(key,val)
      dim likeTemps
      if key<>"" and val<>"" then
	     likeTemps = likesTemp(key,val)
	     if rsWhere<>"" then
		    rsWhere = rsWhere & " or " & likeTemps
		 else
		    rsWhere = likeTemps
		 end if
	  end if
 End Sub

 'order /////////////
 public function orderBy(key,val)
      val = lcase(val)
      if key<>"" and (val="asc" or val="desc") then
		  if rsOrder<>"" then
			 rsOrder = rsOrder & "," & key & " " &val
		  else
			 rsOrder = key & " " &val
		  end if
	  end if
 End function


'---打开数据库记录集 
 public Sub open(db_table,num,paging)
 
      dim rsQuery
	  
      '查询表
      from db_table
      if not paging then selects "*",num
	  '生成查询的SQL
      rsQuery = getSql()
      rs.open rsQuery,conn,1,1
      '是否分页
      if not rs.eof and paging then
	     rs.pagesize= num         '每页记录条数
	     pageSizes  = rs.pagesize
	     allCounts  = rs.recordcount '记录总数
		 
	     pageSizes  = rs.pagesize
	     pageCounts = rs.pagecount 
	     page       = request("page")

	     if not IsNumeric(page) or page="" then
			page = 1
	     else
			page = cint(page)
	     end if
	     if page<1 then
			page=1
	     elseif page>pageCounts then
			page = pageCounts
	     end if
	     rs.absolutepage=Page
	     if page=pageCounts then
			x = allCounts-(pageCounts-1)*pageSizes
	     else
			x = pageSizes
	     end if 
      end if
 End Sub

'---数据库集合移动到下一条
 public function nexts()
    rsNum = rsNum + 1
	rs.movenext
 End function
 
'---判断数据库记录是否已经到最后一条
 public function eof()
	if rsNum>=pageSizes then
	   eof = true
	else
	   if rs.eof then
	      eof = true
	   else
	      eof = false
	   end if  
	end if
 End function
 
'---关闭数据库 
 public function close()
	  rs.close
	  set rs = nothing
 End function
 
 
 
'---返回Sql
 public function getSql()
      dim sqltemp
	  sqltemp = "select " & rsSelect
	  if rsFrom<>"" then sqltemp = sqltemp & " from " & rsFrom
	  if rsWhere<>"" then sqltemp = sqltemp & " where " & rsWhere
	  if rsOrder<>"" then sqltemp = sqltemp & " order by " & rsOrder
      getSql = sqltemp
 End function
 

End Class
%>
