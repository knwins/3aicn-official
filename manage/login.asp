<!--#include file="conn.asp"-->
<html>
<head>
<script language="JavaScript">
<!-->
function check_value(frmlogin)
{
	var userName=frmlogin.userName.value;
	if(userName.length<1)
	{
		alert("帐号不能为空，请您输入帐号名称！");
		frmlogin.userName.focus();
		return false;
	}
	var userPwd=frmlogin.userPwd.value;
	if(userPwd.length<1)
	{
		alert("密码不能为空，请您输入密码！");
		frmlogin.userPwd.focus();
		return false;
	}
	return true;
}
-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link href="back.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.logbtn {
	font-size: 9pt;
	color: #000000;
	height: 14pt;
	width: 36pt;
}
.yy {
	font-size: 12px;
	line-height: normal;
	color: #333333;
	text-decoration: none;
	FILTER: dropshadow(color=#ffffff,offx=1,offy=1); 
}
-->
</style>
</head>

<body bgcolor="E0DFDF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center"><table width="778" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td colspan="3"><img src="images/index1.jpg" width="778" height="125"></td>
        </tr>
        <tr> 
          <td height="14"><img src="images/index2.jpg" width="94" height="135"></td>
          <td width="186" align="center" valign="bottom" background="images/index4.jpg">
<table width="80%" height="65" border="0" cellpadding="0" cellspacing="0">
              <tr>
			<form action="check.asp" method="post" name="frmlogin" id="frmlogin" onSubmit="return check_value(this);">
				  <td align="center"> 
                    <table border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="29%" height="22" align="center" class="yy">用户:</td>
                        <td width="71%" height="15"> <input name="userName" type="text" class="midline"></td>
                      </tr>
                      <tr>
                        <td height="22" align="center" class="yy">密码:</td>
                        <td height="18"><input name="userPwd" type="password" class="midline" ></td>
                      </tr>
                      <tr> 
                        <td height="2" colspan="2"></td>
                      </tr>
                      <tr valign="middle"> 
                        <td height="35" colspan="2">                         
                          <div align="right">
                            <input name="btnmodify" type="image"  id="btnmodify" value="登录" src="images/bottom.jpg">                        
                          </div></td>
                      </tr>
                    </table>
                  </td>
                </form>
              </tr>
            </table></td>
          <td align="right"><img src="images/index3.jpg" width="498" height="135"></td>
        </tr>
        <tr align="right" valign="bottom"> 
          <td height="186" colspan="3" background="images/index5.jpg">&nbsp; </td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
