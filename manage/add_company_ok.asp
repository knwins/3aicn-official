<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%  
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from company"
	rs.open sql,conn,3,2
	rs.addnew
	rs("cname")=trim(request("cname")) '产品名称
	'rs("contenttitle")=trim(request("contenttitle"))
	rs("name")=trim(request("name"))'生成静态页面使用的
	rs("title")=trim(request("title"))
	rs("key")=trim(request("key"))
	rs("description")=trim(request("description"))
	rs("content")=trim(request("content"))
	rs("ver")=session("ver")
	rs.update
%>
<SCRIPT LANGUAGE="JavaScript">
  window.alert("添加成功！");
  window.location.href="company.asp?cname=<%=cname%>";
</SCRIPT>
