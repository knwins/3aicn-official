<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<% 
	dim product_name,orders,content,sql
	product_name = request.form("product_name")
	orders       = request.form("orders")
	content      = request.form("content")
	intro      = request.form("intro")
	types        = request.form("types")
	if request.form("s_image")<>"" then
	s_image        = request.form("s_image")
	else
	s_image=""
	end if

  
  
  set rs=server.createobject("adodb.recordset") 
sql="select * from news"
rs.open sql,conn,1,3 
rs.addnew
rs("PRODUCT_NAME")=PRODUCT_NAME
rs("createdate")=now()
rs("CONTENT")=CONTENT
rs("ORDERS")=ORDERS
rs("intro")=intro
rs("url")=s_image
rs("TYPE")=TYPES
rs("ver")=session("ver")
rs.update
rs.close
set rs=nothing
  
  response.write sql
  conn.execute (sql)
  response.redirect "news_list.asp"
%>
