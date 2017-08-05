<%

'// 获取站点配置信息
function SC(key,filed)
	set myrs = New data
		myrs.where "key","'"&key&"'"
		myrs.open "site_config",1,false
		if not myrs.eof then SC = myrs.rs(filed)
		myrs.close
	set myrs = nothing
end function

'//获取导航信息

%>