<!--#include file="denlu.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<script language=javascript>
function add() {
 
	 if (document.myform.cname.value=="")
	{
	  alert("请输入页面名称！")
	  document.myform.cname.focus()
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
</head>
<body>
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="28"><font color="#FF0000">当前位置 --＞增加企业信息</font>　　　　（带“<font color="#FF0000">*</font>”为必填项） </td>
  </tr>
  <tr>
    <td height="1" bgcolor="#4E6A82"></td>
  </tr>
</table>
<table width="96%" border="0" align="center" cellpadding="2" cellspacing="1">
  <form action="add_company_ok.asp" method="post" name="myform" onSubmit="return add()" >
    <TR>
      <TD HEIGHT="1" COLSPAN="2"></TD>
    </TR>
    <tr>
      <td width="15%" height="28" align="right">栏目名称：</td>
      <td width="85%" height="28"><input name="cname" type="text" id="cname"  size="45" maxlength="50" >
        <font color="#FF0000">*</font></td>
    </tr>
    <tr>
      <td height="28" align="right">栏目文件名：</td>
      <td height="28"><input name="name" type="text" id="name" size="12">
        <font color="#FF0000">*</font></td>
    </tr>
    <tr>
      <td height="28" align="right">网站内容标题：</td>
      <td height="28"><input name="title" type="text" id="title" size="45" maxlength="225"></td>
    </tr>
    <tr>
      <td height="28" align="right">SEO页面关键字：</td>
      <td height="28"><input name="key" type="text" id="key" size="50">
        <font color="#FF0000">*</font></td>
    </tr>
    <tr>
      <td height="28" align="right">SEO页面说明：</td>
      <td height="28"><textarea name="description" cols="40" rows="5" id="description"></textarea></td>
    </tr>
    <tr>
      <td height="28" align="right">页面内容：</td>
      <td height="28"><input type="hidden" name="content" id="content"><iframe id="myiframe" src="edit/ewebeditor.htm?id=content&ReadCookie=0" frameBorder="0" marginHeight="0" marginWidth="0" scrolling="No" width="621" height="459"></iframe></td>
    </tr>
    <tr align="center">
      <td height="28" colspan="2"></td>
    </tr>
    <tr>
      <td height="40" colspan="2" align="center"><input name="image" type="image" src="images/submit.gif" width="72" height="21">
        <img src="images/back.jpg" width="72" height="21" onClick="history.go(-1)"></td>
    </tr>
  </form>
</table>
</body>
</html>
