<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->

<%
xm=trim(request("xm"))
zh=trim(request("zh"))

quanxian=replace(request("quanxian")," ","")
quanxian=","&quanxian&","

conn.execute "update quanxian set xm='"&xm&"',zh='"&zh&"',quanxian='"&quanxian&"' where id="&request("id")
response.write "<script>alert('"&session("xm")&"，你已经修改成功！');window.location.href='quanxian.asp?topage="&request("topage")&"';</script>"
conn.close
%>