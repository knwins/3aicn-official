<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!--#include file="fconn.asp"-->
<%
id=request("id")
types=request("types")
set rs=server.createobject("adodb.recordset")
sql="select * from news where product_id="&id&""
rs.open sql,conn,1,1
product_name=rs("product_name")
content=rs("content")
createdate=rs("createdate")
newsid=rs("product_id")
url=rs("url")
rs.close
set rs=nothing

set rs=server.createobject("adodb.recordset")
sql="select * from news_cate where id="&types&""
rs.open sql,conn,1,1
cname=rs("typename")
filename=rs("filename")
rs.close
set rs=nothing

set rs=server.createobject("adodb.recordset")
sql="select * from site where id=1"
rs.open sql,conn,1,1
title=rs("title")
description=rs("description")
key=rs("key")
rs.close
set rs=nothing%>
<head>
<title><%=product_name%></title>
<meta name="description" content="<%=description%>">
<meta name="keywords" content="<%=key%> - <%=product_name%>">
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
  <div style="margin:0px 0px 8px 0px;"> <img src="images/pic.jpg"/></div>
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
        <div class="boxtitle3"><a><%=product_name%></a><span><a href="/">��ҳ</a> >> <a href="<%=filename%>.html"><%=cname %></a></span></div>
        <div class="boxcont3"><table width="100%" border="0" cellspacing="2" cellpadding="0">
                  <tr>
                    <td height="53" align="center" style="font-size:14px; font-weight:bold"><h1><%=product_name%></h1></td>
                  </tr>
                  <tr>
                    <td align="center">����ʱ�䣺<%=createdate%>&nbsp;&nbsp;&nbsp;�������<script src="count.asp?id=<%=newsid%>" language="javascript"></script>�� &nbsp;&nbsp;&nbsp;������Դ����վ����</td>
                  </tr>
                  <tr>
                    <td class="content"><% if url<>"" then %><p><img src="<%=url%>"/></p><%end if%><p><%=content%></p></td>
                  </tr>
                  <tr>
                    <td height="32" align="right" class="font11"><img src="images/printer.gif"align="absmiddle"> <a href="javascript:window.print()">��ӡ��ҳ</a>&nbsp;&nbsp;<img src="images/close.gif" align="absmiddle"> <a href="javascript:window.close()">�رմ���</a></td>
                  </tr>
                  <tr>
                    <td align="left"><h3 style="font-weight:100; font-size:12px;"><%=description%></h3></td>
                  </tr>
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
      <%=description%></p>
    </div></td>
  </tr>
</table></div>
</div>
</body>
</html>
