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
	  alert("������ҳ�����ƣ�")
	  document.myform.cname.focus()
	  return false
	 } 
	 
	  if (document.myform.title.value=="")
	{
	  alert("������ҳ����⣡")
	  document.myform.title.focus()
	  return false
	 }
	 
if (document.myform.key.value=="")
	{
	  alert("������ҳ��ؼ��֣���")
	  document.myform.key.focus()
	  return false
	 } 
	 if (document.myform.description.value=="")
	{
	  alert("������ҳ��˵������")
	  document.myform.description.focus()
	  return false
	 }  
}
</script>
</head>
<body>
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="28"><font color="#FF0000">��ǰλ�� --��������ҵ��Ϣ</font>��������������<font color="#FF0000">*</font>��Ϊ����� </td>
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
      <td width="15%" height="28" align="right">��Ŀ���ƣ�</td>
      <td width="85%" height="28"><input name="cname" type="text" id="cname"  size="45" maxlength="50" >
        <font color="#FF0000">*</font></td>
    </tr>
    <tr>
      <td height="28" align="right">��Ŀ�ļ�����</td>
      <td height="28"><input name="name" type="text" id="name" size="12">
        <font color="#FF0000">*</font></td>
    </tr>
    <tr>
      <td height="28" align="right">��վ���ݱ��⣺</td>
      <td height="28"><input name="title" type="text" id="title" size="45" maxlength="225"></td>
    </tr>
    <tr>
      <td height="28" align="right">SEOҳ��ؼ��֣�</td>
      <td height="28"><input name="key" type="text" id="key" size="50">
        <font color="#FF0000">*</font></td>
    </tr>
    <tr>
      <td height="28" align="right">SEOҳ��˵����</td>
      <td height="28"><textarea name="description" cols="40" rows="5" id="description"></textarea></td>
    </tr>
    <tr>
      <td height="28" align="right">ҳ�����ݣ�</td>
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