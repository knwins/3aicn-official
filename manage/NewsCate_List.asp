<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%  session("ver")=request("ver")
	set rs=server.createobject("adodb.recordset")
	rs.open "select * from news_cate where ver='"&session("ver")&"' order by listorder asc",conn,1
%>
<html>
<head>
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<style>
body,td,p,select,input{font-size:10pt}
a:link
{
    color:blue;
    font-family: 宋体;
    text-decoration: none
    font-size: 10pt;
}
</style>
<title>新闻类别管理</title>
<script language="javascript" type="text/javascript">
function confirmdelbig()
{
   if(confirm("确定要删除此类吗？此操作不可逆！"))
     return true;
   else
     return false;
}
</script>
</head>
<body>

      <strong><font color="#000000"><img src="../images/indw-3.png" width="10" height="9"> 新闻类别设置</font></strong> <br>
      <br>
      管理选项：<a href="newscate_new.asp">添加新闻大类</a><br>
      <br>
      <table width="780" border="0" cellpadding="0" cellspacing="2" class="border">
        <tr bgcolor="#4296ce" class="title"> 
          <td width="160" height="25" align="center"><strong><font color="#ffffff">栏目名称</font></strong></td>
          <td width="251" height="20" align="center"><strong><font color="#ffffff">操作选项</font></strong></td>
        </tr>
<%
  while not rs.eof 
  	dim space
		cateid = rs("id")
		catename = rs("typename")
		filename=rs("filename")
%>
        
        <tr bgcolor="#eeeeee" class="tdbg"> 
          <td width="160" height="22">
          	<img src="images/tree_folder4.gif" width="15" height="15">
          	<a href="newscate_edit.asp?cateid=<%=cateid%>"><%=catename%></a> <%=filename%>
          </td>
          <td align="right">
							<a href="newscate_edit.asp?cateid=<%=cateid%>">修改</a> 
            | <a href="newscate_edit.asp?mode=Del&cateid=<%=cateid%>" onClick="return confirmdelbig();">删除&nbsp;&nbsp;</a>
          </td>
        </tr>
<%
		rs.movenext
  wend
  rs.close
  set rs=nothing
%>        
        
</table>

</body>
</html>

