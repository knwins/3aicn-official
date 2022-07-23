<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<% 
	cateid  = replace(Request("cateid"),"'","")
	typename_cate=trim(Request("typename_cate"))
	ListOrder=trim(Request("ListOrder"))
	filename=trim(Request("filename"))
	mode=Request("mode")
	set rs=server.CreateObject("ADODB.RecordSet")
	set rschild=server.CreateObject("ADODB.RecordSet")

	if mode="Edit" then
		sql = "update news_cate set typename = '"&typename_cate&"', ListOrder = '"&ListOrder&"', filename = '"&filename&"' "
		sql = sql & "where id = "&cint(cateid)
		Response.Write sql
		conn.Execute (sql)
		Response.Redirect "NewsCate_List.asp?ver="&session("ver")&""
	end if
	
	if mode="Del" then	
			'---------------检查类别下是否有产品-------------------
			rschild.open "select * from news WHERE type= "& cateid &" order by product_id desc" ,conn,1
			if not rschild.eof then
				Response.Write "<script>"
				Response.Write "window.alert('此类别下面有新闻信息，不能删除，请先删除此类别下的新闻信息！');"
				Response.Write "history.back();"
				Response.Write "</script>"
				Response.End 
			else
				sql = "delete * from news_cate where id = "&cateid
				Response.Write sql
		  	conn.Execute (sql)
		  end if
	  	rschild.close
			set rschild=nothing
			Response.Redirect "NewsCate_List.asp?ver="&session("ver")&""		
	end if 
	
	rs.open "select * from news_cate WHERE id= "&cateid &" order by ListOrder desc" ,conn,1
%>
<HTML>
<HEAD><SCRIPT LANGUAGE=javascript>
<!--
	function checkfrom(mode)
	{
		if (mode=="Edit")
		{
			form1.mode.value=mode;
			return true;
		}
		else
		{
			temp=window.confirm("确定要删除这个类吗？")
			if (temp)
			{
				form1.mode.value=mode;
				return true;
			}
			return false;
		}
	}
//-->
</SCRIPT>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
</HEAD>
<BODY>
      <strong><font color="#000000">新闻类别设置</font></strong> <br>
      <br>
      管理选项：<a href="NewsCate_New.asp">添加新闻大类</a><br>
      <br>
			<FORM action="NewsCate_Edit.asp" method=POST name=form1>
    	<table width="414" border="0" cellpadding="0" cellspacing="2" class="border">
        <tr bgcolor="#4296CE" class="title"> 
          <td width="160" height="25" align="center"><strong><font color="#FFFFFF">栏目名称</font></strong></td>
          <td width="251" height="20" align="center"><strong><font color="#FFFFFF">操作选项</font></strong></td>
        </tr>
				<tr bgcolor="#eeeeee" class="tdbg"> 
					<TD width="32%">排序：</TD>
					<TD width="68%"> <INPUT type="text"  name="ListOrder" value="<%=rs("ListOrder")%>" onKeyUp="value=value.replace(/\D+/g,'')"></TD>
				</TR>
	      <tr bgcolor="#eeeeee" class="tdbg"> 
					<TD width=32%>类别名：</TD>
					<TD width="68%"> <INPUT type="text" name="typename_cate" size=40 maxlength=100 value="<%=rs("typename")%>">&nbsp;&nbsp;</TD>
				</TR>
	      <tr bgcolor="#eeeeee" class="tdbg">
	        <TD>类别文件名称:</TD>
	        <TD><INPUT type="text" name="filename" size=40 maxlength=100 value="<%=rs("filename")%>"></TD>
          </TR>
			</TABLE>
			<input type="hidden" name="mode">
			<input type="hidden" name="cateid" value="<%=cateid%>"><p>
		  <INPUT type="submit" value="修 改" name=submit1 onClick="return checkfrom('Edit');"> &nbsp &nbsp
		  <INPUT type="button" value="返 回"  name=button1 onClick="javascript:history.back(-2) ;">
			</FORM>
<%
	rs.close
	set rs=nothing
%>
</body>
</html>
