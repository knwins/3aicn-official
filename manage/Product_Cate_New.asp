<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
	session("loginid")
	parentid=Request("parentid")
	rootid=Request("rootid")
	ProName=trim(Request("ProName"))
	ListOrder=trim(Request("ListOrder"))
	mode=Request("mode")
	
    if mode="add" then
	  	sql = "insert into product_cate(PARENT_ID, ROOT_ID, PRO_CATE_NAME, LISTORDER,ver) "
	  	sql = sql & "values ('"&parentid&"' , '"&rootid&"', '"&ProName&"', '"&ListOrder&"','"&session("ver")&"')"
  
	    Response.Write sql
		conn.Execute (sql)
		Response.Redirect "Product_Cate_List.asp"
	end if
%>

<HTML>
<HEAD><SCRIPT LANGUAGE=javascript>
<!--
	function checkfrom()
	{	var ProName=form1.ProName.value;
		if (ProName=="" || ProName.indexOf("'")>=0)
		{
			window.alert("产品类别名不可为空，不可有特殊字符！");
			return false;
		}
		return true;
	
	}
//-->
</SCRIPT>
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
</HEAD>
<BODY>
<h3 align="center">产品类别设置管理</h3>
<strong><font color="#000000">产品类别设置</font></strong> <br>
<br>
管理选项：<a href="Product_Cate_New.asp?parentid=0&rootid=0">添加产品大类</a><br>
<br>
		<FORM action="Product_Cate_New.asp" method=POST name=form1>
		<table width="414" border="0" cellpadding="0" cellspacing="2" class="border">
    	<tr bgcolor="#4296CE" class="title"> 
        <td width="160" height="25" align="center"><strong><font color="#FFFFFF">序号</font></strong></td>
        <td width="251" height="20" align="center"><strong><font color="#FFFFFF">选项</font></strong></td>
      </tr>
      <tr bgcolor="#eeeeee" class="tdbg"> 
				<TD width="32%">排序：</TD>
				<TD width="68%"> <INPUT type="text" value=0 name="ListOrder" onKeyUp="value=value.replace(/\D+/g,'')"></TD>
			</TR>
      <tr bgcolor="#eeeeee" class="tdbg">  
				<TD width=32%>类别名：</TD>
				<TD width="68%"> <INPUT type="text" name="ProName" size=40 maxlength=100>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
		<input type="hidden" name="mode" value="add">
		<input type="hidden" name="rootid" value="<%=rootid%>">
		<input type="hidden" name="parentid" value="<%=parentid%>"><p>
  	<INPUT type="submit" value="新 增" name=submit1 onClick="return checkfrom();"> &nbsp &nbsp
  	<INPUT type="button" value="返 回"  name=button1 onClick="javascript:history.back() ;">
</FORM>
</BODY>
</HTML>

<%CloseDatabase%>