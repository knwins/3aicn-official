<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
dim rs_pro,sqlpro
sqlpro="select * from site where id="&request("id")
Set rs_pro= Server.CreateObject("ADODB.Recordset")
rs_pro.open sqlpro,conn,1,1
%>
<html>
<head>
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
<body>
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="35"><font color="#FF0000">当前位置： --＞网站参数修改　　</font>（带“<font color="#FF0000">*</font>”为必填项） </td>
  </tr>
</table>
<table width="96%" border="0" align="center" cellpadding="2" cellspacing="1">
  <form action="site_ok.asp" method="post"  name="myform">
    <TR>
      <TD HEIGHT="1" COLSPAN="2"></TD>
    </TR>
    <tr>
      <td width="15%" height="28" align="right">网站名称：</td>
      <td width="85%" height="28"><input name="site" type="text" class="text" id="site" value="<%=rs_pro("site")%>" size="45" maxlength="50"></td>
    </tr>
    <tr>
      <td height="28" align="right">网站首页标题：</td>
      <td height="28"><input name="title" type="text" id="title" value="<%=rs_pro("title")%>" size="60" maxlength="225"></td>
    </tr>
    <tr>
      <td height="28" align="right">网站首页关键字：</td>
      <td height="28"><textarea name="key" cols="40" rows="5" id="key"><%=rs_pro("key")%></textarea></td>
    </tr>
    <tr>
      <td height="28" align="right">网站首页说明：</td>
      <td height="28"><textarea name="description" cols="40" rows="3" id="description"><%=rs_pro("description")%></textarea></td>
    </tr>
    <tr>
      <td height="28" align="right">网 址：</td>
      <td height="28"><input name="url" type="text" id="url" value="<%=rs_pro("url")%>" size="50" maxlength="225"></td>
    </tr>
    <tr>
      <td height="28" align="right">公司地址：</td>
      <td height="28"><input name="add" type="text" id="add" value="<%=rs_pro("add")%>" size="70" maxlength="225"></td>
    </tr>
    <tr>
      <td height="28" align="right">公司电话：</td>
      <td height="28"><input name="companytel" type="text" id="companytel" value="<%=rs_pro("companytel")%>" size="50" maxlength="225"></td>
    </tr>
    <tr>
      <td height="28" align="right">公司传真：</td>
      <td height="28"><input name="fox" type="text" id="fox" value="<%=rs_pro("fox")%>" size="50" maxlength="225"></td>
    </tr>
    <tr>
      <td height="28" align="right">销售热线：</td>
      <td height="28"><input name="tel" type="text" id="tel" value="<%=rs_pro("tel")%>" size="50" maxlength="225"></td>
    </tr>
    <tr>
      <td height="28" align="right">技术支持：</td>
      <td height="28"><input name="support" type="text" id="support" value="<%=rs_pro("support")%>" size="50" maxlength="225"></td>
    </tr>
    <tr>
      <td height="28" align="right">邮箱地址：</td>
      <td height="28"><input name="email" type="text" id="email" value="<%=rs_pro("email")%>" size="50" maxlength="225"></td>
    </tr>
    <tr>
      <td height="40" colspan="2" align="center"><input type="hidden" value="<%=rs_pro("id")%>" name="id">
        <input type="hidden" value="<%=request("topage")%>" name="topage">
        <input name="image2" type="image" src="images/submit.gif" width="72" height="21">
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
