<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
	rootid  = replace(Request("rootid"),"'","")
	parentid  = replace(Request("parentid"),"'","")
	'response.write rootid&"<br>"
	'response.write parentid&"<br>"
	ProName=trim(Request("ProName"))
	ListOrder=trim(Request("ListOrder"))
	mode=Request("mode")
	
	set rs=server.CreateObject("ADODB.RecordSet")
	set rschild=server.CreateObject("ADODB.RecordSet")

	if mode="Edit" then
		sql = "update product_cate set PRO_CATE_NAME = '"&proname&"', ListOrder = '"&ListOrder&"' "
		sql = sql & "where pro_cate_id = "&cint(parentid)
		Response.Write sql
		conn.Execute (sql)
		Response.Redirect "Product_Cate_List.asp"
	end if
	
	if mode="Del" then
		'---------------检查有没有子类别-------------------
		rs.open "select * from product_cate WHERE parent_id= "&parentid &" and root_id= "&parentid &" " ,conn,1
		if not rs.eof then
			rsError = 1
			Response.Write "<script>"
			Response.Write "window.alert('此类别下面有子类，不能删除，请先删除子类别！');"
		 	Response.Write "history.back();"
			Response.Write "</script>"
			Response.End 
			Response.Redirect "Product_Cate_List.asp"
		else
			'---------------检查类别下是否有产品-------------------
			rschild.open "select * from product WHERE type= "& parentid &" order by product_id desc" ,conn,1
			if not rschild.eof then
				Response.Write "<script>"
				Response.Write "window.alert('此类别下面有产品，不能删除，请先删除此类别下的产品！');"
				Response.Write "history.back();"
				Response.Write "</script>"
				Response.End 
			else
				sql = "delete * from product_cate where PRO_CATE_ID = "&parentid
				Response.Write sql
		  	conn.Execute (sql)
		  end if
			rschild.close
			set rschild=nothing
		  '---------------检查类别下是否有产品-------------------
			Response.Redirect "Product_Cate_List.asp"
		end if
		'---------------检查有没有子类别-------------------
	end if 
	
	rs.open "select * from product_cate WHERE pro_cate_id= "&parentid &" order by pro_cate_id desc" ,conn,1
%>
<HTML>
<HEAD>
<SCRIPT LANGUAGE=javascript>
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
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
</HEAD>
<BODY>
<h3 align="center">产品类别设置管理</h3>
<strong><font color="#000000">产品类别设置</font></strong> <br>
<br>
管理选项：<a href="Product_Cate_New.asp?parentid=0&rootid=0">添加产品大类</a><br>
<br>
<FORM action="Product_Cate_Edit.asp" method=POST name=form1>
  <table width="414" border="0" cellpadding="0" cellspacing="2" class="border">
    <tr bgcolor="#4296CE" class="title">
      <td width="160" height="25" align="center"><strong><font color="#FFFFFF">序号</font></strong></td>
      <td width="251" height="20" align="center"><strong><font color="#FFFFFF">选项</font></strong></td>
    </tr>
    <tr bgcolor="#eeeeee" class="tdbg">
      <TD width="32%">排序：</TD>
      <TD width="68%"><INPUT type="text"  name="ListOrder" value="<%=rs("ListOrder")%>" onKeyUp="value=value.replace(/\D+/g,'')"></TD>
    </TR>
    <tr bgcolor="#eeeeee" class="tdbg">
      <TD width=32%>类别名：</TD>
      <TD width="68%"><INPUT type="text" name="ProName" size=40 maxlength=100 value="<%=rs("Pro_Cate_Name")%>">
        &nbsp;&nbsp;</TD>
    </TR>
  </TABLE>
  <input type="hidden" name="mode">
  <input type="hidden" name="parentid" value="<%=parentid%>">
  <p>
    <INPUT type="submit" value="修 改" name=submit1 onClick="return checkfrom('Edit');">
    <INPUT type="button" value="返 回"  name=button1 onClick="javascript:history.back(-2) ;">
</FORM>
<%
	rs.close
	set rs=nothing
%>
</body>
</html>
