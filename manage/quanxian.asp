<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->

<%
if request("Submit2")="ok" and request("delbox")<>"" then
sql="delete from admin where id in ("&request("delbox")&")"
conn.execute(sql)
end if
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<script language="JavaScript" src="js/whir.js"></script>
<style type="text/css">
<!--
.style3 {color: #3664C4}
-->
</style>
</head>
<script language=javascript>
function CheckAll(form)
  {
  for (var i=0;i<form1.elements.length;i++)
    {
    var e = form1.elements[i];
    if (e.name != 'chkall')
       e.checked = form1.chkall.checked;
    }
  }
</script>	
<script language="javascript">
function fy(){
if (isNaN(document.fanye.topage.value)){
alert("请输入数字！")
document.fanye.topage.focus()
return false
}
}
</script>
<style>
BODY {	FONT-SIZE: 12px; scrollbar-face-color:#CFE2FA; scrollbar-shadow-color:#417CC9; scrollbar-highlight-color:#417CC9; scrollbar-3dlight-color:#CFE2FA; scrollbar-darkshadow-color:#CFE2FA; scrollbar-track-color:#CFE2FA; scrollbar-arrow-color:#417CC9}
.style4 {color: #FFFFFF}
</style>
<body text="#000000" leftmargin="0" topmargin="5">
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" height="6">
  <tr> 
    <td width="701" height="25" valign="bottom"><font color="#FF0000">当前位置：权限管理--＞权限管理</font>    </td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">

  <form action="quanxian.asp" method="post" name="form1">
    <tr>
      <td><hr size="1"></td>
    </tr>
    <tr> 
      <td height="35" valign="bottom"> <input type="hidden" name="Submit2">
		<input type="hidden" name="topage" <%if request("topage")<>"" then%>value="<%=request("topage")%>"<%else%>value="1"<%end if%>>
        <input type="image" src="images/del.gif" width="72" height="21" onClick="javascript:if(window.confirm('<%=session("xm")%>，您确实要删除吗？')) window.form1.Submit2.value='ok'">
        &nbsp;<a href="addquanxian.asp"><img src="images/add.gif" width="72" height="21" border="0"></a> 
      </td>
    </tr>
  

  <tr> 
      <td width="100%" align="center" valign="top"> <table width="100%" border="0" cellspacing="1" cellpadding="1" bordercolordark="#ffffff">
          <tr bgcolor="006699"> 
            <td width="3%" height="23" align="center">&nbsp;</td>
            <td width="34%" height="23" align="center"><font color="#FFFFFF">姓名</font></td>
            <td width="18%" height="23" align="center"><font color="#FFFFFF">新闻数据库</font></td>
            <td width="17%" align="center" class="style4">产品展示数据库</td>
            <td width="15%" align="center" class="style4">权限管理</td>
            <td width="13%" height="23" align="center" class="style4"><font color="#FFFFFF">相关操作</font></td>
          </tr>
<%
set rs=server.createobject("adodb.recordset")
aaa="select * from admin where id is not null order by id desc"
rs.open aaa,conn,3
%>
          <% if rs.bof and rs.eof then %>
          <tr bgcolor="#F6F6F6"> 
            <td height="22" colspan="8" align=center>数据库里还没有权限!</td>
          </tr>
          <% else %>
          <%
rs.pagesize=25
if not isempty(request("topage")) then
topage=cint(request("topage"))
if topage>rs.pagecount then
rs.absolutepage=rs.pagecount
elseif topage<=0 then
rs.absolutepage=1
else
rs.absolutepage=topage
end if
end if
intcurpage=rs.absolutepage

%>
          <%
				for i=1 to rs.pagesize
				if rs.eof then
				exit for
				end if
				%>
          <tr bgcolor="#F6F6F6" onMouseOver="this.bgColor='#e6e6e6';" onMouseOut="this.bgColor='#F6F6F6';"> 
            <td height="22" align="center">

<%if rs("id")>1 then%><input type="checkbox" name="delbox" value="<%=rs("id")%>"><%end if%></td>
            <td align="center"><%=rs("xm")%></td>

<td height="22" align="center"><%if instr(1,rs("quanxian"),",1,")>0 then%><font color="#ff0000">√</font><%else%>×<%end if%></td>
<td height="22" align="center"><%if instr(1,rs("quanxian"),",2,")>0 then%>
  <font color="#ff0000">√</font>
  <%else%>
  ×
  <%end if%></td>
<td height="22" align="center"><%if instr(1,rs("quanxian"),",3,")>0 then%>
  <font color="#ff0000">√</font>
  <%else%>
  ×
  <%end if%></td>
<td height="22" align="center"><a href="quanxian_modi.asp?id=<%=rs("id")%>&topage=<%=intcurpage%>">修改</a></td>
</tr>
          <%
				  rs.movenext
				  next 
				  end if %>
        </table></td>
    </tr>
    <tr> 
      <td width="100%" height="30">　　 
        <input type="checkbox" name="chkall" value="on" onClick="CheckAll(this.form)"> 
        <span class="style3">全部选中</span> </td>
    </tr>
  </form>
</table>
<%if intcurpage<>0 then%>
<table width="98%" height="26" border="0" align="center" cellpadding="0" cellspacing="0">
  <form action=quanxian.asp method=post name="fanye" onSubmit="return(fy())">
    <tr>
      <td height="26" align="right">目前在第<%=intcurpage %>页，共有<%=rs.pagecount%>页,共有<%=rs.recordcount%>条记录　
          <%if intcurpage<>1 then%>
          <a href=quanxian.asp?topage=<%=1%>>第一页</a> <a href=quanxian.asp?topage=<%=intcurpage-1%>>上一页</a>
          <%
		end if
		if intcurpage<>rs.pagecount then
		%>
          <a href=quanxian.asp?topage=<%=intcurpage+1%>>下一页</a> <a href=quanxian.asp?topage=<%=rs.pagecount%>>最后一页</a>
          <%end if%>
        请输入页次:
        <input type=text size=1 name=topage value="<%=intcurpage%>" class="text">
        <input type="submit" name="Submit" value="Go" class="text">
&nbsp;&nbsp;　　 </td>
    </tr>
  </form>
</table>
<%end if%>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
%>