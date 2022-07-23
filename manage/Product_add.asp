<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
	<% 
	dim PRODUCT_NAME, ORDERS, CONTENT, INTRO, COMMEND, S_IMAGE, B_IMAGE, sql
	PRODUCT_NAME = Request.Form("PRODUCT_NAME")
	ORDERS       = Request.Form("ORDERS")
	INTRO        = Request.Form("INTRO")
	CONTENT      = Request.Form("content")
	COMMEND      = Request.Form("COMMEND")
	S_IMAGE      = Request.Form("s_image")
	B_IMAGE      = Request.Form("B_IMAGE")' ´óÍ¼Æ¬
	TYPES        = Request.Form("TYPES")
	
	if INTRO = "" then
	INTRO = 0
	end if
	if COMMEND = "checkbox" then 
	COMMEND = 1
	else 
	COMMEND = 0
	end if
	set rss=server.createobject("adodb.recordset") 
	sqls="select * from product_cate where pro_cate_id="&TYPES&""
	rss.open sqls,conn,1,3 
	bigtype=rss("root_id")


set rs=server.createobject("adodb.recordset") 
sql="select * from product"
rs.open sql,conn,1,3 
rs.addnew
rs("PRODUCT_NAME")=PRODUCT_NAME
rs("COMMEND")=COMMEND
rs("CONTENT")=CONTENT
rs("ORDERS")=ORDERS
rs("S_IMAGE")=S_IMAGE
rs("INTRO")=INTRO
rs("B_IMAGE")=B_IMAGE
rs("TYPE")=TYPES
rs("BIGTYPE")=bigtype
rs("ver")=session("ver")
rs.update
rs.close
set rs=nothing
	Response.Write sql
	conn.Execute (sql)
	Response.Redirect "Product_List.asp"
	%>

<%CloseDatabase%>


