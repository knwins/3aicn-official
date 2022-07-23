<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
if session("ver")="" or  request("ver")="cn" or request("ver")="en" then
      session("ver")=request("ver")
	end if
	%>
<html>
<head>
<title>留言</title>
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
</head>
<body>
<br>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#B3DD99">
  <tr>
    <td width="100%" valign="top" bgcolor="#FFFFFF" ><br>
	 <% 
	page=clng(request("page"))		 
	set rs=server.createobject("adodb.recordset") 
	sql="select * from [msg] where ver='"&session("ver")&"' order by id desc"
	rs.open sql,conn,1,1
    if rs.eof and rs.bof then
     response.write("暂时没有记录")
     else 
	rs.pagesize=3
	if page=0 then page=1 
	pages=rs.pagecount
	if page > pages then page=pages
	rs.absolutepage=page  
	for j=1 to rs.pagesize 
%>
<br>
<table width="96%" border="0" align="center" cellspacing="1" bgcolor="#D5EAC6" style="padding-top:3px;">
                              <tr valign="top">
                                <td height="24" colspan="3" valign="middle" bgcolor="#FFFFFF" style="padding-left:5px;"><span style="font-size:13px;"><b>主题：</b>
                                  <% response.Write rs("title")%>
                                </span>
                                  <span style="font-size:12px; color:#999999"> 时间:
                                  <% response.Write rs("date")%>
                                </span><a href="user_info.asp?userid=<%=rs("userid")%>"> <b>所属用户</b></a></td>
                              </tr>
                              <tr valign="top">
                                <td colspan="3" valign="top" bgcolor="#F4FAF1" style="padding:5px;"><b>内容：</b>
                                <% response.Write rs("msg")%></td>
                              </tr>
							   <form name="myform" method="post" action="msg_save.asp?id=<%=rs("id")%>&page=<%=page%>"><tr valign="top">
                                <td width="12%" valign="top" bgcolor="#E8F4E1" style="padding:10px"><b>回复：</b></td>
							    <td width="59%" align="left" valign="middle" bgcolor="#E8F4E1" style="padding:10px">
                                 <textarea name="sendmsg" cols="60" rows="3" id="sendmsg"><%=rs("sendmsg")%></textarea>                                </td>
							    <td width="29%" align="left" valign="middle" bgcolor="#E8F4E1" style="padding:10px"><input name="image" type="image" src="images/modi.gif" width="72" height="21">&nbsp;&nbsp;<a href="msg_del.asp?id=<%=rs("id")%>&page=<%=page%>"><img src="images/del.gif" width="72" height="21" border="0"></a></td>
							   </tr></form>
      </table> <%
rs.movenext
if rs.eof then exit for
next
%>
      <br></td>
  </tr>
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF" style="padding:3px;">
<form method=post action="msg.asp">
<%if page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=msg.asp?page=1>首页</a>&nbsp;"
    response.write "<a href=msg.asp?page=" & page-1 & ">上一页</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=msg.asp?page=" & (page+1) & ">"
    response.write "下一页</a> <a href=msg.asp?page="&rs.pagecount&">尾页</a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&page&"</font>/"&rs.pagecount&"</strong>页 "
    response.write "&nbsp;共<b><font color='#ff0000'>"&rs.recordcount&"</font></b>条记录 <b>"&rs.pagesize&"</b>条记录/页"
   response.write " 转到：<input type='text' name='page' size=4 maxlength=10 class=input value="&page&">"
   response.write " <input class=input type='submit'  value='Go'  name='cndok'></span></p>"     
%>
    </form>	<% 
end if
rs.close
set rs=nothing
%></td>
  </tr>
</table>

</body>
</html>

