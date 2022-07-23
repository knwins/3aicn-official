<!--#include file="conn.asp"-->
<%
userName=trim(request("userName"))
userPwd=trim(request("userPwd"))
set rs=server.createobject("adodb.recordset")
sql="select * from admin where zh='"&userName&"'"
rs.open sql,conn,3
if Not Rs.Eof then
if userPwd<>trim(rs("pass")) then
response.write "<script>alert('你的用户名或密码有错误！');window.history.back(-1)</script>"
response.end
else
session.timeout=60
session("xm")=rs("xm")
session("zh")=rs("zh")
session("userlogin")="ok"
response.redirect "all.asp"
end if
else
response.write "<script>alert('你的用户名或密码有错误！');window.history.back(-1)</script>"
end if
%>