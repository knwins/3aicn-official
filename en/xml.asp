<!--#include file="fconn.asp"-->
<%
response.Write"<?xml version='1.0' encoding='utf-8'?>"
response.Write "<bcaster autoPlayTime='5'>"
set rs=server.createobject("adodb.recordset")
rs.open"select top 3 * from news where ver='en' and type=35 order by orders desc",conn,1  
while not rs.eof
response.Write "<item item_url='/"
response.Write rs("url")
response.Write "' link='"
response.Write rs("product_id")
response.Write ".html'>"
response.Write "</item>"
rs.movenext
wend 
rs.close
set rs=nothing
response.Write "</bcaster>"
%>
