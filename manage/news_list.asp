<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
if session("ver")="" or  request("ver")="cn" or request("ver")="en" then
      session("ver")=request("ver")
	end if
	
	keyname=trim(request("keyname"))
	keytype=trim(request("keytype"))
	
if request("Submit2")="ok" and request("delbox")<>"" then
	conn.execute "delete from news where product_id in ("&request("delbox")&")"
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
    <td width="701" height="20" valign="middle"><font color="#FF0000">当前位置：<%=request("name")%>&gt;&gt; 信息管理</font></td>
  </tr>
</table>

  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1">
  <form action="news_list.asp" method="post" name="form">
    <tr align="center">
      <td width="30%" height="20" align="center"><div align="right">关键字搜索：
          <input name="keyname" type="text" id="name" value="<%=request("keyname")%>" size="14">
        </div></td>
      <td width="70%" align="center"><div align="left"> &nbsp;<input name="image" type="image" src="images/shousuo.gif" width="72" height="21">
        </div></td>
    </tr>
	</form>
	<form action="news_list.asp" method="post" name="form">
    <tr align="center">
      <td height="20" align="right">请输入类别查询：
      <select  name="keytype"><%call showtype()%></select></td>
      <td align="left">&nbsp;<input name="image" type="image" src="images/shousuo.gif" width="72" height="21"></td>
    </tr>
	</form>
</table>

<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="1" bgcolor="#4E6A82"></td>
  </tr>
  <form action="news_list.asp" method="post" name="form1">
    <tr>
      <td height="35" valign="bottom"><input type="hidden" name="Submit2">
        <input type="hidden" name="keyname" value="<%=request("keyname")%>">
		<input type="hidden" name="keytype" value="<%=request("keytype")%>">
        <input type="hidden" name="topage" <%if request("topage")<>"" then%>value="<%=request("topage")%>"<%else%>value="1"<%end if%>>
        <input type="image" src="images/del.gif" width="72" height="21" onClick="Delok(this.form)">
        &nbsp;<a href="add_news.asp?keytype=<%=keytype%>"><img src="images/add.gif" width="72" height="21" border="0"></a> </td>
    </tr>
    <tr>
      <td width="100%" align="center" valign="top"><table width="100%" border="0" cellspacing="1" cellpadding="1" bordercolordark="#ffffff">
          <tr bgcolor="006699">
            <td width="9%" align="center">&nbsp;</td>
            <td width="11%" height="23" align="center"><font color="#FFFFFF">序号</font></td>
            <td width="48%" height="23" align="center"><font color="#FFFFFF">新闻标题</font></td>
            <td width="14%" height="23" align="center"><b><font color="#FFFFFF">类 别</font></b></td>
            <td width="18%" height="23" colspan="2" align="center"><font color="#FFFFFF">相关操作</font></td>
          </tr>
          <%
set rs=server.createobject("adodb.recordset")
if keyname<>"" then
    rs.open "select * from news where ver='"&session("ver")&"' and product_name like '%"&keyname&"%' order by product_id desc",conn,3
else if keytype<>"" then
    rs.open "select * from news where ver='"&session("ver")&"' and type="&keytype&" order by product_id desc",conn,3
else
  rs.open "select * from news where ver='"&session("ver")&"' order by product_id desc",conn,3
end if
end if
typeid = rs("type")
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

      	set rscate=server.createobject("adodb.recordset")
				rscate.open "select * from news_cate where id= "&typeid &" and ver='"&session("ver")&"' order by listorder asc" ,conn,1
				catename=rscate("typename")
				set rscate=nothing
				'-------------------------------
%>
          <%
				for i=1 to rs.pagesize
                if rs.eof then
				exit for
				end if
				%>
          <tr bgcolor="#F6F6F6" onMouseOver="this.bgColor='#e6e6e6';" onMouseOut="this.bgColor='#F6F6F6';">
            <td align="center"><input type="checkbox" name="delbox" value="<%=rs("product_id")%>"></td>
            <td height="22" align="center"><%=i+(topage-1)*(rs.pagesize)%></td>
            <td ><%=rs("product_name")%></td>
            <td ><%=catename%></td>
            <td width="18%" height="22" align="center"><a href="news_modi.asp?id=<%=rs("product_id")%>&topage=<%=intcurpage%>&typeids=<%=rs("type")%>">修改</a></td>
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
  <form action="news_list.asp" method="post" name="fanye" onSubmit="return(fy())">
    <tr>
      <td height="26" align="right">目前在第<%=intcurpage %>页，共有<%=rs.pagecount%>页,共有<%=rs.recordcount%>条记录　
        <%if intcurpage<>1 then%>
        <a href="news_list.asp?topage=<%=1%>&keytype=<%=keytype%>">第一页</a> <a href="news_list.asp?topage=<%=intcurpage-1%>&keytype=<%=keytype%>">上一页</a>
        <%
		end if
		if intcurpage<>rs.pagecount then
		%>
        <a href="news_list.asp?topage=<%=intcurpage+1%>&keytype=<%=keytype%>">下一页</a> <a href="news_list.asp?topage=<%=rs.pagecount%>&keytype=<%=keytype%>">最后一页</a>
        <%end if%>
        请输入页次:
        <input type=text size=3 name=topage value="<%=intcurpage%>" class="text">
        &nbsp;&nbsp;
        <input type="submit" name="Submit" value="Go" class="text">
        <input name="cname" type="hidden" id="keyname" value="<%=request("keyname")%>">
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
<%
sub getcate(cateid)
	set rscate=server.createobject("adodb.recordset")
	rscate.open "select * from news_cate where id= "&cateid &" and ver='"&session("ver")&"' order by listorder asc" ,conn,1
	catename=rscate("typename")
	set rscate=nothing
end sub
sub showtype()
	set rscate=server.createobject("adodb.recordset")
	rscate.open "select * from news_cate where ver='"&session("ver")&"'order by listorder asc",conn,1
	while not rscate.eof
		str1="<option value="""& rscate("id") & """>" & rscate("typename") & "</option>"
		response.write str1
		rscate.movenext
	wend 
	set rscate=nothing
end sub
sub	getchiled(cateid,level)
		set rschiled=objcate.rsgetproductcate("ado","chiled",cstr(cateid))
		while not rschiled.eof
			objcate.datafromado(rschiled)
			if objcate.productcateid =search2 then
				str1="<option value="""& objcate.productcateid &  """selected>" 
			else
				str1="<option value="""& objcate.productcateid &  """>" 
			end if
			call list(str1,level)
			call getchiled(objcate.productcateid,(level+1))
			rschiled.movenext
		wend
		rschiled.close
		set rschiled=nothing
end sub
sub	list(str,level)
		dim space
		for i=1 to level
			space=space+""
		next
		str2=objcate.typename & "</option>"
		str=str & space & str2
		response.write str
end sub

sub setpublish
	dim strsqltext
	if not intswilldele="" then
		response.write "intswilldele==="&intswilldele
		obj.product_id = intswilldele
		obj.save "ado","publish"
	end if
end sub
%>