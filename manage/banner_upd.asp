<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<% 
	dim id,title,p_image,url,sql
	id = request.form("id")
	title = request.form("title")
	p_image = request.form("p_image")
	url = request.form("url")
	
	sql = "update banner set title = '"&title&"',p_image = '"&p_image&"',"
	sql = sql & "url = '"&url&"'"
	sql = sql & " where id ="&id&""
	response.write sql
	conn.execute (sql)
	response.write "更新成功！"
	response.write "<script>window.opener.location.reload();</script>"
	response.redirect "banner_list.asp"
%>

