<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->

<%
set rs=server.createobject("adodb.recordset")
rs.open "select * from admin where id="&request("id"),conn,3

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<script language="JavaScript" src="js/whir.js"></script>
<title>无标题文档</title>
<style type="text/css">
<!--
.style2 {color: #3664C4}
.style3 {color: #FF0000}
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
	 
r = document.form1.quanxian;
var j=0;
for(i=0;i<r.length;i++){
  if(r[i].checked)j++;
}
if(j<1){
alert("权限最少要选中其中一个！");
return false;
} 
}
</script>
<body leftmargin="0" topmargin="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="35"><font color="#FF0000">当前位置：权限管理--＞修改权限　</font></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <form action="quanxian_modiok.asp" method="post" name="form1" onSubmit="return aa()">
    <TR> 
      <TD HEIGHT="1" COLSPAN="3" BGCOLOR="#4E6A82"></TD>
    </TR>
    <tr> 
      <td height="50" colspan="3" align="center"><span class="style2">姓　　名：
        <input name="xm" type="text" class="text" id="xm" value="<%=rs("xm")%>" size="30" maxlength="50">
        <span class="style3">* </span></span></td>
    </tr>
    <tr>
      <td height="50" colspan="3" align="center"><span class="style2">帐　　号：
          <input name="zh" type="text" class="text" id="zh" value="<%=rs("zh")%>" size="30" maxlength="25" readonly>
          <span class="style3">      * </span></span></td>
    </tr>
    <tr align="center">
      <td height="50" colspan="3"><div align="right" class="style3">
        <div align="center">栏目权限:<span class="style2"><input name="quanxian" type="checkbox" id="quanxian" value="1" <%if instr(1,rs("quanxian"),",1,")>0 then%>checked<%end if%>></span><font color="#1D579E">新闻数据库</font><span class="style2">　
            <input name="quanxian" type="checkbox" id="quanxian" value="2" <%if instr(1,rs("quanxian"),",2,")>0 then%>checked<%end if%>>
            <font color="#1D579E">产品展示数据库</font></span><span class="style2">　
<input name="quanxian" type="checkbox" id="quanxian" value="3" <%if instr(1,rs("quanxian"),",3,")>0 then%>checked<%end if%>>
<font color="#1D579E">权限管理</font></span></div>
      </div>        </td>
    </tr>
    <tr> 
      <td height="45" colspan="3" align="center"> <input type="hidden" value="<%=rs("id")%>" name="id">
        <input type="hidden" value="<%=request("topage")%>" name="topage">
      <input type="image" src="images/modi.gif" width="72" height="21">　
      <img src="images/back.jpg" width="72" height="21" onClick="history.go(-1)"></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
%>