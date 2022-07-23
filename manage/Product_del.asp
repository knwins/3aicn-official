<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->

<%
  dim sql  
  sql = "delete * from product where PRODUCT_ID = "&Request.QueryString("id")
  conn.Execute (sql)
  Response.Redirect "Product_List.asp"
%>
<%CloseDatabase%>