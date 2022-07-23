<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
'conn.execute "update login set down_date='"&now()&"' where id="&session("kid")
session.Abandon()
response.write "<script>window.parent.location.href='login.asp'</script>"
%>