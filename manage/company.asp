<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
if session("ver")="" or  request("ver")="cn" or request("ver")="en" then
      session("ver")=request("ver")
	end if
if request("Submit2")="ok" and request("delbox")<>"" then
	conn.execute "delete from company where id in ("&request("delbox")&")"
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<script language="JavaScript" src="js/whir.js"></script>
<style type="text/css">
<!--
.style1 {color: #FFFFFF}
.style3 {color: #4e6a82}
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

function Delok(fm)
{
  var checkbool
  checkbool=false
  for(var i=0;i<fm.elements.length;i++)
   {
     var e=fm.elements[i];
	 //alert(e)
     if(e.checked)//显示选中的记录
      {
	     checkbool=true;
	  }
   }
	
	if (checkbool)
	{
	  if(confirm("确认要删除吗?"))
  		 {
 		   window.form1.Submit2.value='ok';
   		 }	
	}
	else
	{
	  alert("请选择要删除的项!")
	  return false;
	}
}
</script>
<body text="#000000" leftmargin="0" topmargin="5">
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" height="25">
  <tr>
    <td width="701" height="20" valign="middle"><font color="#FF0000">当前位置：<%=request("name")%>&gt;&gt; 企业信息管理</font></td>
  </tr>
</table>
<form action="company.asp" method="post" name="form">
  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr align="center">
      <td width="65%" height="20" align="center"><div align="right">关键字搜索：
          <input name="cname" type="text" class="text" id="name" value="<%=request("name")%>" size="25">
          <input type="hidden" value="eee" name="Submit">
        </div></td>
      <td width="35%" align="center"><div align="left"> &nbsp;
          <input name="image" type="image" src="images/shousuo.gif" width="72" height="21">
        </div></td>
    </tr>
  </table>
</form>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="1" bgcolor="#4E6A82"></td>
  </tr>
  <form action="company.asp" method="post" name="form1">
    <tr>
      <td height="35" valign="bottom"><input type="hidden" name="Submit2">
        <input type="hidden" name="nb" value="<%=request("cname")%>">
        <input type="hidden" name="key" value="<%=request("name")%>">
        <input type="hidden" name="topage" <%if request("topage")<>"" then%>value="<%=request("topage")%>"<%else%>value="1"<%end if%>>
        <input type="image" src="images/del.gif" width="72" height="21" onClick="Delok(this.form)">
        &nbsp;<a href="add_company.asp"><img src="images/add.gif" width="72" height="21" border="0"></a> </td>
    </tr>
    <tr>
      <td width="100%" align="center" valign="top"><table width="100%" border="0" cellspacing="1" cellpadding="1" bordercolordark="#ffffff">
          <tr bgcolor="006699">
            <td width="5%" align="center">&nbsp;</td>
            <td width="7%" height="23" align="center"><font color="#FFFFFF">序号</font></td>
            <td width="25%" height="23" align="center"><font color="#FFFFFF">页面名称</font></td>
            <!--<td width="19%" height="23" align="center"><font color="#FFFFFF">页面文件名</font></td> -->
            <td width="32%" height="23" align="center"><font color="#FFFFFF">页面标题</font></td>
            <td width="12%" height="23" colspan="2" align="center"><font color="#FFFFFF">相关操作</font></td>
          </tr>
          <%
set rs=server.createobject("adodb.recordset")
if request("Submit")<>"" then
     typenames=request("cname")
    rs.open "select * from company where ver='"&session("ver")&"' and cname like '%"&typenames&"%' order by id desc",conn,3
else
  rs.open "select * from company where ver='"&session("ver")&"' order by id desc",conn,3
end if
%>
          <% if rs.bof and rs.eof then %>
          <tr bgcolor="#F6F6F6">
            <td height="22" colspan="12" align=center>数据库里还没有数据!</td>
          </tr>
          <% else %>
          <%
if request("topage")="" then
topage=1
end if
rs.pagesize=16
if not isempty(request("topage")) then
topage=clng(request("topage"))
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
            <td align="center"><input type="checkbox" name="delbox" value="<%=rs("id")%>"></td>
            <td height="22" align="center"><%=i+(topage-1)*(rs.pagesize)%></td>
            <td ><div align="left"><%=rs("cname")%></div></td>
            <!--<td ><div align="left"><%'=rs("name")%></div></td> -->
            <td ><div align="left"><%=rs("title")%></div></td>
            <td width="12%" height="22" align="center"><a href="company_modi.asp?id=<%=rs("id")%>&topage=<%=intcurpage%>">修改</a></td>
          </tr>
          <%
				  rs.movenext
				  next
				  end if

		  %>
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
  <form action="company.asp" method="post" name="fanye" onSubmit="return(fy())">
    <tr>
      <td height="26" align="right">目前在第<%=intcurpage %>页，共有<%=rs.pagecount%>页,共有<%=rs.recordcount%>条记录　
        <%if intcurpage<>1 then%>
        <a href="company.asp?topage=<%=1%>&Submit=eee">第一页</a> <a href="company.asp?topage=<%=intcurpage-1%>&Submit=eee">上一页</a>
        <%
		end if
		if intcurpage<>rs.pagecount then
		%>
        <a href="company.asp?topage=<%=intcurpage+1%>&Submit=eee">下一页</a> <a href="company.asp?topage=<%=rs.pagecount%>&Submit=eee">最后一页</a>
        <%end if%>
        请输入页次:
        <input type=text size=3 name=topage value="<%=intcurpage%>" class="text">
        &nbsp;&nbsp;
        <input type="submit" name="Submit" value="Go" class="text">
        <input name="cname" type="hidden" id="cname" value="<%=request("cname")%>">
        <input type="hidden" value="eee" name="Submit">
      </td>
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
