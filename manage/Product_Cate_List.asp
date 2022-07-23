<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
if session("ver")="" or  request("ver")="cn" or request("ver")="en" then
      session("ver")=request("ver")
	end if
	set rs=server.CreateObject("ADODB.RecordSet")
	rs.open "select * from product_cate WHERE ver='"&session("ver")&"' and ROOT_ID=0 order by listorder asc",conn,1
%>
<HTML>
<HEAD>
<TITLE>产品类别管理</TITLE>
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<script language="JavaScript" type="text/JavaScript">
function ConfirmDelBig()
{
   if(confirm("确定要删除此类吗？此操作不可逆！"))
     return true;
   else
     return false;
}
</script>
</HEAD>
<BODY>
			<h3 align="center">产品类别设置管理</h3>
      <strong><font color="#000000">产品类别设置</font></strong> <br>
      <br>
      管理选项：<a href="Product_Cate_New.asp?parentid=0&rootid=0">添加产品大类</a><br>
      <br>
      <table width="414" border="0" cellpadding="0" cellspacing="2" class="border">
        <tr bgcolor="#4296CE" class="title"> 
          <td width="160" height="25" align="center"><strong><font color="#FFFFFF">栏目名称</font></strong></td>
          <td width="251" height="20" align="center"><strong><font color="#FFFFFF">操作选项</font></strong></td>
        </tr>
<%
  while not rs.eof 
  	dim space
		parentid = rs("PRO_CATE_ID")
		catename = rs("PRO_CATE_NAME")
%>
        
        <tr bgcolor="#eeeeee" class="tdbg"> 
          <td width="160" height="22">
          	<img src="images/o.gif" width="16" height="16">
          	<A HREF="Product_Cate_Edit.asp?parentid=<%=parentid%>"><b><%=catename%></b></a>
          </td>
          <td align="right">
          		<A HREF='Product_Cate_New.asp?parentid=<%=parentid%>&catename=<%=catename%>&rootid=<%=parentid%>'>添加中类</a> 
            | <A HREF="Product_Cate_Edit.asp?parentid=<%=parentid%>">修改</a> 
            | <a href="Product_Cate_Edit.asp?mode=Del&rootid=<%=rs("Parent_ID")%>&parentid=<%=parentid%>" onClick="return ConfirmDelBig();">删除</a>
           &nbsp;&nbsp;</td>
        </tr>
			<%
				'response.write "parentid========"&parentid
				set rschild=server.CreateObject("ADODB.RecordSet")
				rschild.open "select * from product_cate where parent_id="&parentid&" order by listorder asc",conn,1
				while not rschild.eof
			%>
        <tr bgcolor="#eeeeee" class="tdbg"> 
          <td width="160" height="22">
          	&nbsp;&nbsp;<img src="images/big.gif" width="16" height="16">
          	<A HREF="Product_Cate_Edit.asp?parentid=<%=rschild("pro_cate_id")%>"><%=rschild("pro_cate_name")%></a></td>
          <td align="right">
            	<!--<A HREF='Product_Cate_New.asp?rootid=<%=parentid%>&catename=<%=catename%>&parentid=<%=rschild("pro_cate_id")%>'>添加小类</a>  -->
            | <a href="Product_Cate_Edit.asp?parentid=<%=rschild("pro_cate_id")%>">修改</a> 
            | <a href="Product_Cate_Edit.asp?mode=Del&rootid=<%=rschild("Parent_ID")%>&parentid=<%=rschild("pro_cate_id")%>" onClick="return ConfirmDelBig();">删除</a>
       </td>
        </tr>
		<% 
		set rsmall=server.CreateObject("ADODB.RecordSet")
				rsmall.open "select * from product_cate where parent_id="&rschild("pro_cate_id")&" and root_id="&rs("PRO_CATE_ID")&" and ver='"&session("ver")&"' order by listorder asc",conn,1
				while not rsmall.eof
				%>
		        <tr bgcolor="#eeeeee" class="tdbg"> 
          <td width="160" height="22">
          	&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/small.gif" width="16" height="16">
          	<A HREF="Product_Cate_Edit.asp?parentid=<%=rsmall("pro_cate_id")%>"><%=rsmall("pro_cate_name")%></a>
          </td>
          <td align="right">
            	<a href="Product_Cate_Edit.asp?parentid=<%=rsmall("pro_cate_id")%>">修改</a> 
            | <a href="Product_Cate_Edit.asp?mode=Del&rootid=<%=rsmall("Parent_ID")%>&parentid=<%=rsmall("pro_cate_id")%>" onClick="return ConfirmDelBig();">删除</a>
          </td>
        </tr>
					<% rsmall.movenext
						wend 
						rsmall.close
						set rsmall=nothing
					rschild.MoveNext
				wend 
				rschild.close
				set rschild=nothing
			%>
<%
		rs.movenext
  wend
  rs.close
  set rs=nothing
%>        
        
    </table>

</BODY>
</HTML>
<%
sub GetChiled(parentid,level,rootid)
	set rschiled=server.CreateObject("ADODB.RecordSet")
	rschiled.open "select * from product_cate where parent_id="&parentid&" order by listorder asc ",conn,1
	while not rschiled.eof
		str1="<A HREF=""Product_Cate_Edit.asp?parentid=" & rschiled("PRO_CATE_ID") & """>" & rschiled("PRO_CATE_NAME") & "(序号："& rschiled("listorder") &")</A>---------"
		str1=str1 & "  <A HREF=""Product_Cate_New.asp?parentid=" & rschiled("PRO_CATE_ID") & "&catename="& rschiled("PRO_CATE_ID") & "&rootid="& rootid & """><font color=red>+</font></A>"
		str1=str1 & "<br>"
		call list(str1,level)
		call Getchiled(rschiled("PRO_CATE_ID"),(level+1),rootid)
	rschiled.movenext
	wend
	rschiled.close
	set rschiled=nothing
end sub

sub list(str,level)
	space=""
	for i=1 to level
		space=space+""
	next
	str=space & str
	Response.Write str
end sub
%>
<%CloseDatabase%>