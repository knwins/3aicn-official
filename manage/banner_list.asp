<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%  if session("ver")="" or  request("ver")="cn" or request("ver")="en" then
      session("ver")=request("ver")
	end if
	bymode=trim(request("bymode"))
	search=trim(request("search"))
	search2=trim(request("search2"))

	dim rs,rec,mc,ms,xh
	set rs=server.createobject("adodb.recordset")
	
  if trim(bymode)="byname" then
		rs.open "select * from banner where ver='"&session("ver")&"' and  name like '%" + search + "%' order by id desc",conn,1
	else
		rs.open "select * from banner where ver='"&session("ver")&"' order by id desc",conn,1
	end if
	
	'-------------------------------------------
	pagesize=12
	icount=rs.recordcount 
	
	if (icount mod pagesize)=0 then
		pagecount=int(icount/pagesize)
	else
		pagecount=int(icount/pagesize)+1
	end if
	if request("currentpage")<>"" and request("currentpage")<>"0" then
		currentpage=cint(request("currentpage"))
	else
		currentpage=1
	end if
	if not rs.eof then
		rs.pagesize =pagesize
		rs.absolutepage=currentpage
		ipage=0
	end if
	if  icount=0 then
		currentpage=0
	end if
	'-------------------------------------------

%>
<html>
<head>
<title>后台数据管理中心</title>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<script language=javascript>
<!--
	function submitform(mode){
		form1.action="banner_list.asp?bymode="+mode;
		form1.submit();
	}
  function del(did){
	    temp=window.confirm("确定要删除吗？")
	    if( temp ){
	      location.href = "banner_del.asp?id="+did;
	    }
	}
//-->

</script>
</head>
<body bgcolor="#ffffff" leftmargin=0 topmargin=0>
<table width="95%" cellspacing="0"  border=0 cellpadding="2" >
<tr>
	<td align=center>
		<table width="95%" border="1" cellspacing="0" cellpadding="2" bordercolor="#cccccc">
		  <tr bgcolor="#e4e4e4"> 
		    <td align="center"> 
		      <table width="100%" border="0" cellspacing="0" cellpadding="0">
		      <form name="form1" method="post">
		        <tr> 
		          <td width="20%">
		          	请输入新闻标题模糊查询：
		            <input name=search size=51 value="<%=search%>">
		            &nbsp;<a href="javascript:submitform('byname')">查询</a>		          </td>
		        </tr>
		      </form>
		      </table>
		    </td>
		  </tr>
		</table>
	</td>
</tr>

<tr>
	<td>
		<table width="95%" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr> 
		    <td>&nbsp;</td>
		  </tr>
		  <tr>
		  	<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
					  <tr>
					    <td>共 <%=icount%> 条 第<%=currentpage%>页（共<%=pagecount%>页） | 
					      <%if currentpage>1 then
					      		response.write "<a href=banner_list.asp?search=" & search & "&search2=" & search2 & "&bymode=" & bymode & "&currentpage=1>首页</a>&nbsp;&nbsp;"
										response.write "<a href=banner_list.asp?search=" & search & "&search2=" & search2 & "&bymode=" & bymode & "&currentpage=" & (currentpage-1) & ">上一页</a>&nbsp;&nbsp;"
									  else
									  response.write "首页&nbsp;&nbsp;"
										response.write "上一页&nbsp;&nbsp;"
									end if	
								%>
					      <%if currentpage<pagecount then
										response.write "<a href=banner_list.asp?search=" & search & "&search2=" & search2 & "&bymode=" & bymode & "&currentpage=" & (currentpage+1) & "> 下一页</a>"
										response.write "<a href=banner_list.asp?search=" & search & "&search2=" & search2 & "&bymode=" & bymode & "&currentpage=" & pagecount & "> 末页</a>"
									  else
										response.write " 下一页&nbsp;&nbsp;"
										response.write " 末页"
									end if
								%>
								| <a href='banner_new.asp'>新增</a>
					    </td>
					  </tr>
					</table>
		      <table width="100%" cellpadding="1" bgcolor="#ffffff" cellspacing="1" bordercolor="#cccccc">
					  <tr bgcolor="#e4e4e4"> 
					    <td  height="23"align=center bgcolor="006699"><b ><font color="#FFFFFF">序号</font></b></td>
					    <td align=center bgcolor="006699"><b><font color="#FFFFFF">标题</font></b></td>
					    <td align=center bgcolor="006699"><b><font color="#FFFFFF">URL</font></b></td>
					    <td align=center bgcolor="006699"><b><font color="#FFFFFF">操作</font></b></td>
					  </tr>
<%   
	while not rs.eof and ipage<pagesize
%>
		        <tr onMouseOver="this.bgColor='#e6e6e6';" onMouseOut="this.bgColor='#F6F6F6';"> 
		          <td height="22"><a href="banner_update.asp?id=<%=rs("id")%>"><%=rs("id")%></a></td>
		          <td align="left">
		            <a href="banner_update.asp?id=<%=rs("id")%>&page=<%=CurrentPage%>"><%=rs("title")%></a>		          </td>
		          <td><%=rs("url")%></td>
		          <td><a href="javascript:del(<%=rs("id")%>)">删除</a></td>
		        </tr>                          
<%                          
		rs.movenext
		ipage=ipage+1
	wend 
	rs.close
	set rs=nothing
%>
		      </table>
		    </td>
			</tr>       
		</table>
	</td>
</tr>
</table>
</body>       
</html>