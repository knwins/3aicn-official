<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<% 
dim title,ver,p_image,url,sql
title = request.form("title")
p_image = request.form("p_image")
url = request.form("url")
sql = "insert into banner(title,ver,p_image,url)values('"&title&"','"&session("ver")&"','"&p_image&"','"&url&"')"
conn.execute (sql)
response.redirect "banner_list.asp"
%>