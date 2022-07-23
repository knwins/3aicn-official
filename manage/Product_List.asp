<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
if session("ver")="" or  request("ver")="cn" or request("ver")="en" then
      session("ver")=request("ver")
	end if
	
	typeid=request.QueryString("typeid")
	bymode=trim(Request("bymode"))
	search=trim(Request("Search"))
	search2=trim(Request("Search2"))

	dim rs,rec,mc,ms,xh
	set rs=server.CreateObject("ADODB.RecordSet")
	
  	if trim(bymode)="byname" then
		rs.open "select * from product where ver='"&session("ver")&"' and product_name like '%" + search + "%' order by ORDERS asc",conn,1
	elseif trim(bymode)="bytype" then
  	rs.open "select * from product where ver='"&session("ver")&"' and type = " + search2 + " order by ORDERS asc",conn,1
	elseif typeid<>"" then
	   rs.open "select * from product where ver='"&session("ver")&"' and  TYPE="&typeid&" order by ORDERS asc",conn,1
	else 
		rs.open "select * from product where ver='"&session("ver")&"' order by ORDERS asc",conn,1
	end if
	
	'-------------------------------------------
	pagesize=35
	iCount=rs.RecordCount 
	
	if (iCount mod pagesize)=0 then
		pagecount=int(iCount/pagesize)
	else
		pagecount=int(iCount/pagesize)+1
	end if
	if Request("currentpage")<>"" and Request("currentpage")<>"0" then
		currentpage=Cint(Request("currentpage"))
	else
		currentpage=1
	end if
	if not rs.EOF then
		rs.PageSize =pagesize
		rs.AbsolutePage=currentpage
		ipage=0
	end if
	if  iCount=0 then
		currentpage=0
	end if
	'-------------------------------------------

%>
<html>
<head>
<title>后台数据管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<SCRIPT LANGUAGE=javascript>
<!--
	function submitForm(mode){
		form1.action="Product_List.asp?bymode="+mode;
		form1.submit();
	}
  function del(did){
	    temp=window.confirm("确定要删除吗？")
	    if( temp ){
	      location.href = "Product_del.asp?id="+did;
	    }
	}
//-->

</SCRIPT>
</head>
<body bgcolor="#FFFFFF" leftmargin=0 topmargin=0>
<table width="98%"  border=0 align="center" cellpadding="2" cellspacing="0" >
<tr>
	<td align=center>
		<table width="100%" border="1" cellspacing="0" cellpadding="2" bordercolor="#cccccc">
		  <tr bgcolor="#e4e4e4"> 
		    <td align="center"> 
		      <table width="100%" border="0" cellspacing="0" cellpadding="0">
		      <form name="form1" method="post">
		        <tr> 
		          <td width="20%">
		          	请输入“产品名称”模糊查询：
		            <input name=search size=51 value="<%=search%>">
		            &nbsp;<a href="javascript:submitForm('byname')">查询</a>
		          </td>
		        </tr>
		        <tr>
		          <td width="20%">
		          请输入类别查询：&nbsp;&nbsp;&nbsp;&nbsp;
                	<SELECT  name="search2"><%call showtype()%></SELECT>
		            &nbsp;<a href="javascript:submitForm('bytype')">查询</a>
		          </td>
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
		<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr> 
		    <td>&nbsp;</td>
		  </tr>
		  <tr>
		  	<td>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
					  <tr>
					    <td>共 <%=iCount%> 条 第<%=currentpage%>页（共<%=pagecount%>页） | 
					      <%if currentpage>1 then
					      		Response.Write "<A href=Product_List.asp?search=" & search & "&search2=" & search2 & "&bymode=" & bymode & "&currentpage=1>首页</A>&nbsp;&nbsp;"
										Response.Write "<A href=Product_List.asp?search=" & search & "&search2=" & search2 & "&bymode=" & bymode & "&currentpage=" & (currentpage-1) & ">上一页</A>&nbsp;&nbsp;"
									  else
									  Response.Write "首页&nbsp;&nbsp;"
										Response.Write "上一页&nbsp;&nbsp;"
									end if	
								%>
					      <%if currentpage<pagecount then
										Response.Write "<A href=Product_List.asp?search=" & search & "&search2=" & search2 & "&bymode=" & bymode & "&currentpage=" & (currentpage+1) & "> 下一页</A>"
										Response.Write "<A href=Product_List.asp?search=" & search & "&search2=" & search2 & "&bymode=" & bymode & "&currentpage=" & pagecount & "> 末页</A>"
									  else
										Response.Write " 下一页&nbsp;&nbsp;"
										Response.Write " 末页"
									end if
								%>
								| <a href='Product_new.asp'>新增</a>
					    </td>
					  </tr>
					</table>
		      <table width="100%" border="0" cellspacing="1" cellpadding="1" bordercolordark="#ffffff">
					  <tr bgcolor="#e4e4e4"> 
					    <td width="7%" height="23" align=center bgcolor="006699"><font color="#FFFFFF">编号</font></td>
					    <td width="28%" align=center bgcolor="006699"><font color="#FFFFFF">产品名称</font></td>
					    <td width="19%" align=center bgcolor="#006699"><font color="#FFFFFF">小图片</font></td>
					    <td width="26%" align=center bgcolor="006699"><font color="#FFFFFF">类别</font></td>
					    <td width="10%" align=center bgcolor="#006699"><font color="#FFFFFF">是否推荐</font></td>
					    <td width="10%" align=center bgcolor="006699"><font color="#FFFFFF">操作</font></td>
					  </tr>
<%   
	while not rs.eof and ipage<pagesize
      	call GetCate(rs("type"))  
      	'-------------------------------
      	cateid = rs("type")
      	set rscate=server.CreateObject("ADODB.RecordSet")
				rscate.open "select * from product_cate WHERE pro_cate_id= "&cateid &"" ,conn,1
				catename=rscate("pro_cate_name")
				TYPEBIG=rscate("ROOT_ID")
				set RS_BIG=server.CreateObject("ADODB.RecordSet")
				RS_BIG.open "select * from product_cate WHERE pro_cate_id= "&TYPEBIG &"" ,conn,1
				TYPEBIGNAME=RS_BIG("pro_cate_name")
				
				
				
				
				
				set rscate=nothing
				'-------------------------------
%>
		        <tr bgcolor="#F6F6F6" onMouseOver="this.bgColor='#e6e6e6';" onMouseOut="this.bgColor='#F6F6F6';"> 
		          <td><%=rs("ORDERS")%></td>
		          <td align="left">
		            <a href="Product_update.asp?id=<%=rs("product_id")%>"><%=rs("product_name")%></a>
		          </td>
		          <td><a href="../<%=rs("S_IMAGE")%>" target="_blank"><img src="../<%=rs("S_IMAGE")%>" width="29" height="28" border="0"></a></td>
		          <td><%=TYPEBIGNAME%>>><%=catename%></td>
				  <td><%if rs("COMMEND")=0 then response.Write "否" else response.Write "是" end if%></td>
		          <td>
		          	<a href="Product_update.asp?id=<%=rs("product_id")%>&page=<%=CurrentPage%>">修改</a>
		          	<a href="javascript:del(<%=rs("PRODUCT_ID")%>)">删除</a>
		          </td>
		        </tr>                          
<%                          
		rs.MoveNext
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

<%
sub GetCate(cateid)
	set rscate=server.CreateObject("ADODB.RecordSet")
	rscate.open "select * from product_cate WHERE pro_cate_id= "&cateid &" order by listorder asc" ,conn,1
	catename=rscate("pro_cate_name")
	rscate.close
	set rscate=nothing
end sub
sub showtype()
	set rscate=server.CreateObject("ADODB.RecordSet")
	rscate.open "select * from product_cate WHERE ver='"&session("ver")&"' and ROOT_ID=0 order by listorder ASC",conn,1
	while not rscate.eof
		str1="<option value="""& rscate("pro_cate_id") & """>" & rscate("Pro_Cate_Name") & "</option>"
		Response.Write str1
		call Getchiled(rscate("pro_cate_id"),1) 
		rscate.movenext
	wend 
	set rscate=nothing
end sub
sub	GetChiled(cateid,level)
		set rschiled=server.CreateObject("ADODB.RecordSet")
		rschiled.open "select * from product_cate WHERE ver='"&session("ver")&"' and ROOT_ID="&cateid&" order by listorder ASC",conn,1
		while not rschiled.eof
			if rschiled("pro_cate_id") =search2 then
				str1="<option value="""& rschiled("pro_cate_id") &  """selected>" 
			else
				str1="<option value="""& rschiled("pro_cate_id") &  """>" 
			end if
			call list(rschiled("Pro_Cate_Name"), str1,level)
			call Getchiled(rschiled("pro_cate_id"),(level+1))
			rschiled.movenext
		wend
		set rschiled=nothing
end sub
sub	list(catename, str,level)
		dim space
		for i=1 to level
			space=space+""
		next
		str2=catename & "</option>"
		str=str & space & str2
		Response.Write str
end sub

sub setpublish
	dim strSQLText
	if not intsWillDele="" then
		Response.Write "intsWillDele==="&intsWillDele
		obj.Product_ID = intsWillDele
		obj.Save "ADO","PUBLISH"
	end if
end sub
%>