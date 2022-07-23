<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<script language="JavaScript" src="whir.js"></script>
<title>无标题文档</title>
<style type="text/css">
<!--
.style3 {color: #FF0000}
.style2 {color: #3664C4}
.style4 {color: #000000}
-->
</style>
</head>

<script language=javascript>
function aa() {
if (document.form1.xm.value=="")
	{
	  alert("请输入姓名！")
	  document.form1.xm.focus()
	  return false
	 }	
if (document.form1.zh.value=="")
	{
	  alert("请输入帐号！")
	  document.form1.zh.focus()
	  return false
	 }
	 
if (document.form1.pass.value=="")
	{
	  alert("请输入密码！")
	  document.form1.pass.focus()
	  return false
	 }	 	 	 	 
}
</script>	 
<body leftmargin="0" topmargin="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="35"><font color="#FF0000">当前位置：权限管理--＞增加管理员　</font>　<span class="style4">（带“</span><font color="#FF0000">*</font><span class="style4">”为必填项）</span> </td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <form action="addquanxianok.asp" method="post" name="form1" onSubmit="return aa()">
    <TR> 
      <TD HEIGHT="1" COLSPAN="2" BGCOLOR="#4E6A82"></TD>
    </TR>
    <tr>
      <td height="30" colspan="2" align="center">&nbsp;</td>
    </tr>
    <tr>
      <td width="38%" height="30" align="center"><div align="right"><font color="#1D579E">姓 名：</font>
        </div></td>
      <td width="62%" height="30" align="center"><div align="left">
        <input name="xm" type="text" id="xm" size="20" maxlength="20" class="text">
        <span class="style3">*</span></div></td>
    </tr>
    <tr> 
      <td height="30" align="center"><div align="right"><font color="#1D579E">帐 号：</font> </div></td>
      <td height="30" align="center"><div align="left">
        <input name="zh" type="text" id="zh" size="20" maxlength="20" class="text">
        <span class="style3">*</span></div></td>
    </tr>
    <tr> 
      <td height="30" align="center"><div align="right"><font color="#1D579E">密 码：</font></div></td>
      <td height="30" align="center"><div align="left">
        <input name="pass" type="password" id="pass" size="20" maxlength="20" class="text">
      <span class="style3">*</span>        <span class="style3"></span></div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2" align="center">&nbsp;</td>
    </tr>
    <tr> 
      <td height="50" colspan="2" align="center"> <div align="center"><span class="style3">栏目权限:</span><input name="quanxian" type="checkbox" id="quanxian" value="1">
          <font color="#1D579E">新闻数据库</font><font color="#1D579E">　
          <input name="quanxian" type="checkbox" id="quanxian" value="2">
          产品展示数据库
          </font><font color="#1D579E">　
          <input name="quanxian" type="checkbox" id="quanxian" value="3">
          <span class="style2">权限管理</span></font></div>
      </td>
    </tr>
    <tr> 
      <td height="45" colspan="2" align="center"> <input type="image" src="images/submit.gif" name="Submit" value=" 提 交 " onClick="return(aa())">　
      <img src="images/back.jpg" width="72" height="21" onClick="history.go(-1)"></td>
    </tr>
  </form>
</table>
</body>
</html>
