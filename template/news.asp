<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!--#include file="fconn.asp"-->
<%
types=request("type")
set rs=server.createobject("adodb.recordset")
sql="select * from site where id=1"
rs.open sql,conn,1,1
title=rs("title")
description=rs("description")
key=rs("key")
rs.close
set rs=nothing

set rs=server.createobject("adodb.recordset")
sql="select * from news_cate where id="&types&""
rs.open sql,conn,1,1
cname=rs("typename")
newsname=rs("filename")  
rs.close
set rs=nothing

%>
<head>
<title><%=title%></title>
<meta name="Keywords" content="<%=key%>" />
<meta name="Description" content="<%=description%>" />
<meta http-equiv="Content-type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>
<body>
<div id="container">
  <div id="header" style="height:140px;">
    <div class="logo"><img src="images/logo.gif"></div>
    <h1 style="position:absolute; top:76px; left:24px; color:#000; font-weight:normal; font-size:12px; margin:0; padding:0;"> <%=description%></h1>
    <div class="language" style="top:15px;">
      <p> <a href="/" target="_self">���İ�</a> <a href="/en">ENGLISH</a></p>
    </div>	<div class="tel">�ͷ���4000442113</div>
    <div id="nav" style="top:100px;">
      <ul><li style="padding-left:10px;"><a href="index.html">��&nbsp;&nbsp;ҳ</a></li>
        <li><a href="probig_78.html">3M�������</a></li>
        <li><a href="solutions.html">�����豸</a></li>
		 <li><a href="applications.html">��ƷӦ��</a></li>
        <li><a href="caselist.html">���̰���</a></li>
        <li><a href="news.html">���¶�̬</a></li>
        <li><a href="aboutus.html">��������</a></li>
        <li><a href="contactus.html">��ϵ����</a></li>

      </ul>
    </div>
  </div>
  <div style="margin:0px 0px 8px 0px;"> <img src="images/pic.jpg" /></div>
  <div id="maincontent">
  <table width="950" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div class="boxe_com">
      <div class="box">
        <div class="boxtitle"><a>����������ϵ��ʽ</a></div>
        <div class="boxcont"><br />
          <p>��˾��ַ�������и����������빤ҵ��526��618<br />
            �ͷ���4000442113
<br />
            ���棺+86-755-28260069
<br />
            ���ʣ�goumike@163.com<br />
          �ͷ�24Сʱ���ߣ�13600442113</p>
          <p>��ַ��<a href="Http://www.3aicn.com">Http://www.3aicn.com</a></p>
          <p><img src="images/contactus.gif" width="212" height="134"></p>
        </div>
      </div>
      <div class="box2_com">
        <div class="boxtitle3"><a ><%=cname %></a><span><a href="/">��ҳ</a> >> <a href="<%=newsname%>.html"><%=cname %></a></span></div>
        <div class="boxcont3" style="padding-top:12px;">
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
          <% 
	page=clng(request("page"))		 
	set rs=server.createobject("adodb.recordset") 
	sql="select * from news where type="&types&" and ver='cn'"
	rs.open sql,conn,1,1
    if rs.eof and rs.bof then
     response.write("��ʱû�м�¼")
     else 
	rs.pagesize=12
	if page=0 then page=1 
	pages=rs.pagecount
	if page > pages then page=pages
	rs.absolutepage=page  
	for j=1 to rs.pagesize 
%>
          <tr>
            <td width="6%" height="24" align="center" >��</td>
            <td height="24" style="border-bottom: #999999 1px dotted"><a href="<%=rs("product_id")%>.html"><%= rs("product_name") %></a> <% if rs("url")<>"" then response.Write " <font color=#FF0000> ͼ</font>" end if %>��<font color="#999999">[<%=year(rs("createdate"))%>-<%=month(rs("createdate"))%>-<%= day(rs("createdate")) %>]</font></td>
          </tr>
          <%
rs.movenext
if rs.eof then exit for
next
%>
          <tr valign="bottom">
            <td height="30" colspan="2" align="center" ><table border="0" cellspacing="0" cellpadding="0" class="font">
                          <tr>
                            <td align="center" height="22">
							<div id="pageClass3">
                                <ul>
                                  <% for k=1 to rs.pagecount%>
                                  <li class="currentpage"><a href="<%=newsname%>_<%=k%>.html"><%=k%></a></li>
                                  <% next %>
                                </ul>
                            </div></td>
                          </tr>
                    </table></td>
          </tr>
          <% 
end if
rs.close
set rs=nothing
%>
        </table>
		</div>
      </div>
    </div></td>
  </tr>
  <tr>
    <td><div id="footer">
      <div class="footernav"> <a href="/index.html">��ҳ</a>| <a href="product.html">��Ʒ����</a>| <a href="caselist.html">���̰���</a>| <a href="rongyu.html">��ҵ����</a>| <a href="qanda.html">��������</a>| <a href="aboutus.html">��������</a>| <a href="contactus.html">��ϵ��ʽ</a> </div>
      <p style=" padding-top:5px;">Copyright &copy; 2003-2011 ���������������������޹�˾, All Rights Reserved &nbsp;&nbsp; ��ICP��06023278�� <br />
        �㶫���ڣ� �����и����������빤ҵ��526��618  �ͷ���4000442113
 ���棺0755-28260069
 �ֻ���13600442113<br />
      <%=description%> <script src="http://s16.cnzz.com/stat.php?id=3224905&web_id=3224905&show=pic" language="JavaScript"></script></p>
    </div></td>
  </tr>
</table></div>
</div>
</body>
</html>
