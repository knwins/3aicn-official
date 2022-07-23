<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
id=request("id")
set rs=server.CreateObject("adodb.recordset")
sql="select * from company where id="&id
rs.open sql,conn,3,2
    types=trim(request("typename")) '定义大类名称
	'rs("contenttitle")=trim(request("contenttitle"))
    stype=trim(request("stypename")) '定义大类名称    
	rs("cname")=trim(request("cname")) '产品名称
	rs("name")=trim(request("name"))
	rs("title")=trim(request("title"))
	rs("key")=trim(request("key"))
	rs("description")=trim(request("description"))
	rs("content")=trim(request("content"))
rs.update
%>
<SCRIPT LANGUAGE="JavaScript">

  window.alert("已成功修改该记录！");
  window.location.href="company.asp?topage=<%=request("topage")%>";

</SCRIPT>