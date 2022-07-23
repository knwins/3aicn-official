<!--#include file="fconn.asp"-->
<%
id=request("id")
sql = "update news set hot=hot+1 where product_id="&id&""
conn.execute (sql)
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select * from news where product_id="&id&""  
rs.open sql,conn,1,1
hot=rs("hot")
%>
document.write(<%=hot%>)

