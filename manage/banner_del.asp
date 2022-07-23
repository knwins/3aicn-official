<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->


<%
  dim sql  
  sql = "delete * from banner where id = "&Request.QueryString("id")
  conn.Execute (sql)
  Response.Redirect "banner_List.asp"
%>

<%CloseDatabase%>
