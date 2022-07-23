<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
dim rs_pro,sqlpro
sqlpro="select * from company where id="&request("id")
Set rs_pro= Server.CreateObject("ADODB.Recordset")
rs_pro.open sqlpro,conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
</head>
<script language=javascript>
function add() {
 
	 if (document.myform.cname.value=="")
	{
	  alert("请输入页面名称！")
	  document.myform.cname.focus()
	  return false
	 } 
	 if (document.myform.name.value=="")
	{
	  alert("请输入页面文件名称！")
	  document.myform.name.focus()
	  return false
	 }  	
	  if (document.myform.title.value=="")
	{
	  alert("请输入页面标题！")
	  document.myform.title.focus()
	  return false
	 }
	 
if (document.myform.key.value=="")
	{
	  alert("请输入页面关键字！！")
	  document.myform.key.focus()
	  return false
	 } 
	 if (document.myform.description.value=="")
	{
	  alert("请输入页面说明！！")
	  document.myform.description.focus()
	  return false
	 }  
}
</script>
<body>
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="35"><font color="#FF0000">当前位置：<%=request("nbmc")%> --＞企业信息管理　　</font>（带“<font color="#FF0000">*</font>”为必填项） </td>
  </tr>
</table>
<table width="96%" border="0" align="center" cellpadding="2" cellspacing="1">
  <form action="company_modi_ok.asp" method="post" name="myform" onSubmit="return add()" >
    <TR>
      <TD HEIGHT="1" COLSPAN="2"></TD>
    </TR>
    <tr>
      <td width="15%" height="28" align="right">栏目名称：</td>
      <td width="85%" height="28"><input name="cname" type="text" id="cname" value="<%=rs_pro("cname")%>"  size="45" maxlength="50" >
          <font color="#FF0000">*</font></td>
    </tr>
    <tr>
      <td height="28" align="right">栏目文件名：</td>
      <td height="28"><input name="name" type="text" id="name" value="<%=rs_pro("name")%>" size="12">
          <font color="#FF0000">*</font></td>
    </tr>
    <tr>
      <td height="28" align="right">网站内容标题：</td>
      <td height="28"><input name="title" type="text" id="title" value="<%=rs_pro("title")%>" size="40" maxlength="100"></td>
    </tr>
    <tr>
      <td height="28" align="right">SEO页面关键字：</td>
      <td height="28"><input name="key" type="text" id="key" value="<%=rs_pro("key")%>" size="50">
          <font color="#FF0000">*</font></td>
    </tr>
    <tr>
      <td height="28" align="right">SEO页面说明：</td>
      <td height="28"><textarea name="description" cols="50" rows="3" id="description"><%=rs_pro("description")%></textarea></td>
    </tr>
    <tr>
      <td height="28" align="right">页面内容：</td>
      <td height="28"><textarea name="content" id="content" style="display:'none'"><%=rs_pro("content")%></textarea><iframe id="myiframe" src="edit/ewebeditor.htm?id=content&ReadCookie=0" frameBorder="0" marginHeight="0" marginWidth="0" scrolling="No" width="621" height="459"></iframe></td>
    </tr>
    <tr align="center">
      <td height="28" colspan="2"></td>
    </tr>
    <tr>
      <td height="40" colspan="2" align="center"><input type="hidden" value="<%=rs_pro("id")%>" name="id"><input type="hidden" value="<%=request("topage")%>" name="topage"><input name="image" type="image" src="images/submit.gif" width="72" height="21">
          <img src="images/back.jpg" width="72" height="21" onClick="history.go(-1)"></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
rs_pro.close
set rs_pro=nothing
conn.close
%>
