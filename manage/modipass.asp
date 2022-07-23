<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->

<%
set rs=server.createobject("adodb.recordset")
rs.open "select pass from admin where zh='"&session("zh")&"' ",conn,3
pass=rs(0)
rs.close

if request("newpass")<>"" then
  conn.execute "update admin set pass='"&request("newpass")&"' where zh='"&session("zh")&"' "
  response.write "<script>alert('"&session("xm")&"，你的密码已经修改成功！');window.location.href='face.asp'</script>"
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<title>无标题文档</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.style1 {color: #FFFFFF}
-->
</style></head>

<script language=javascript>
function aa(){
if (document.all.item("oldpass").value=="")
	{
	  alert("请输入旧密码！")
	  document.all.item("oldpass").focus()
	  return false
	 }
	 
if (document.all.item("oldpass").value!="<%=pass%>")
	{
	  alert("输入的旧密码不正确！")
	  document.all.item("oldpass").focus()
	  return false
	 }
	 
if (document.all.item("newpass").value=="")
	{
	  alert("请输入新密码！")
	  document.all.item("newpass").focus()
	  return false
	 }
	 
if (document.all.item("newpass1").value!=document.all.item("newpass").value)
	{
	  alert("两次输入的新密码不一样！")
	  document.all.item("newpass1").focus()
	  return false
	 }	  
}
</script>
<body>
<table width="100%" height="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="top"><table width="96%" border="0" cellspacing="0" cellpadding="0" align="center" height="6">
      <tr>
        <td width="701" height="35" valign="middle"><font color="#FF0000">当前位置：密码修改</font></td>
      </tr>
      <tr> 
      <td height="1" bgcolor="#000000"></td>
  </tr>
</table>
      <br>
      <br>
      <br>
      <br>
      <br>
      <br>
      <br>
      <table width="500" border="0">
	<form action="modipass.asp" method="post" name="form1" onSubmit="return aa()">
      <tr align="center" bgcolor="006699">
        <td height="25" colspan="2"><span class="style1">修改密码</span></td>
      </tr>
      <tr bgcolor="#F6F6F6">
        <td height="25" align="right" bgcolor="#F6F6F6">旧　密　码：</td>
        <td><input name="oldpass" type="password" class="text" id="oldpass" size="20"></td>
      </tr>
      <tr bgcolor="#F6F6F6">
        <td height="25" align="right" bgcolor="#F6F6F6">新　密　码：</td>
        <td height="25"><input name="newpass" type="password" class="text" id="newpass" size="20"></td>
      </tr>
      <tr bgcolor="#F6F6F6">
        <td height="25" align="right" bgcolor="#F6F6F6">确认新密码：</td>
        <td height="25"><input name="newpass1" type="password" class="text" id="newpass1" size="20"></td>
      </tr>
      <tr align="center" bgcolor="#F6F6F6">
        <td height="35" colspan="2"><input name="imageField" type="image" src="images/submit.gif" width="72" height="21" border="0"></td>
      </tr>
	  </form>
    </table></td>
  </tr>
</table>
</body>
</html>
<%
set rs=nothing
conn.close
%>