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
sql="select * from site where id=2"
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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>
<body>
<div id="container">
  <div id="header" style="height:140px;">
    <div class="logo"><img src="images/logo.gif"></div>
    <h1 style="position:absolute; top:76px; left:24px; color:#000; font-weight:normal; font-size:12px; margin:0; padding:0;"> <%=description%></h1>
    <div class="language" style="top:15px;">
      <p> <a href="/" target="_self">China</a> <a href="/en">ENGLISH</a></p>
    </div>
	<div class="tel">Service Hotline:4000442113 </div>
    <div id="nav" style="top:100px;">
      <ul><li style="padding-left:10px;"><a href="index.html">Home</a></li>
        <li><a href="probig_137.html">3M Fire Protection Materials</a></li>
        <li><a href="solutions.html">Power Equipment</a></li>
		 <li><a href="applications.html">Application</a></li>
        <li><a href="caselist.html">Case</a></li>
        <li><a href="news.html">News</a></li>
        <li><a href="aboutus.html">About Us</a></li>
        <li><a href="contactus.html">Contact Us</a></li>

      </ul>
    </div>
  </div>
  <div style="margin:0px 0px 8px 0px;"> <img src="images/pic.jpg"/></div>
  <div id="maincontent">
  <table width="950" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div class="boxe_com">
      <div class="box">
        <div class="boxtitle"><a>Contact Us</a></div>
        <div class="boxcont" style="padding-top:9px;">
          <p><b>Shenzhen Company</b></p>
          <p>Address:Futian District, Shenzhen Industrial Park 526 618 Bagualing<br />
          Service Tel:4000442113<br />Fax:+86-755-28260069<br />E-Mail:goumike@163.com<br />Mobile:13600442113</p>
          <p>Website:<a href="Http://www.3aicn.com">Http://www.3aicn.com</a></p>
          <p><img src="images/contactus.gif" width="212" height="123"></p>
        </div>
      </div>
      <div class="box2_com">
        <div class="boxtitle3"><a><%=product_name%></a><span><a href="/">Home</a> >> <a href="<%=filename%>.html"><%=cname %></a></span></div>
        <div class="boxcont3"><table width="100%" border="0" cellspacing="2" cellpadding="0">
                  <tr>
                    <td height="53" align="center" style="font-size:14px; font-weight:bold"><%=product_name%></td>
                  </tr>
                  <tr>
                    <td align="center">Time<%=createdate%>&nbsp;&nbsp;&nbsp;Hits<script src="count.asp?id=<%=newsid%>" language="javascript"></script> &nbsp;&nbsp;&nbsp;From:3aicn</td>
                  </tr>
                  <tr>
                    <td class="content"><% if url<>"" then %><p><img src="/<%=url%>"/></p><%end if%><p><%=content%></p></td>
                  </tr>
                  <tr>
                    <td height="32" align="right" class="font11"><img src="images/printer.gif"align="absmiddle"> <a href="javascript:window.print()">Print</a>&nbsp;&nbsp;<img src="images/close.gif" align="absmiddle"> <a href="javascript:window.close()">Close</a></td>
                  </tr>
              </table> 
        </div>
      </div>
    </div></td>
  </tr>
  <tr>
    <td><div id="footer">
      <div class="footernav"> <a href="/index.html">home</a>| <a href="probig_137.html">product</a>| <a href="caselist.html">case</a>| <a href="rongyu.html">Honor</a>| <a href="qanda.html">asked questions</a>| <a href="aboutus.html">about us</a>| <a href="contactus.html">contact us</a> </div>
      <p style=" padding-top:5px;">Copyright &copy; 2003-2011 Shenzhen Sanai Power Technology Co.,Ltd, All Rights Reserved &nbsp;&nbsp; <br />
        Shenzhen Guangdong Address: Futian District, Shenzhen Industrial Park 526 618 Bagualing  Service tel:4000442113  FAX:0755-28260069  Mobile:13600442113 <br>
<%=description%><!-- JiaThis Button BEGIN -->

<!-- JiaThis Button END --> </p>
    </div></td>
  </tr>
</table></div>
</div>
</body>
</html>
