<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!--#include file="fconn.asp"-->
<%
id=request("id")
set rs=server.CreateObject("adodb.recordset")
sql="select * from company where id="&id&""
rs.open sql,conn,1,1
contents=rs("content")
title=rs("title")
content=rs("content")
cname=rs("cname")
descriptions=rs("description")
key=rs("key")
rs.close
set rs=nothing%>
<head>
<title><%=title%></title>
<meta http-equiv="Content-type" content="text/html; charset=gb2312" />
<meta name="description" content="<%=descriptions%>">
<meta name="keywords" content="<%=key%>">
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>
<body>
<div id="container">
  <div id="header" style="height:140px;">
    <div class="logo"><img src="images/logo.gif"></div>
    <h1 style="position:absolute; top:76px; left:24px; color:#000; font-weight:normal; font-size:12px; margin:0; padding:0;"> <%=descriptions%></h1>
    <div class="language" style="top:15px;">
      <p> <a href="/" target="_self">���İ�</a> <a href="/en">ENGLISH</a></p>
    </div>	
    <div class="tel">�ͷ���400442113</div>
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
  <div style="margin:0px 0px 4px 0px;"> <img src="images/banner02.jpg" width="950" height="150" /></div>
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
          <p><img src="images/contactus.gif"></p>
        </div>
      </div>
      <div class="box2_com">
        <div class="boxtitle3"><a ><%=cname %></a><span><a href="/">��ҳ</a> >> <%=cname %></span></div>
        <div class="content" style="padding-top:12px;"><%=content%> </div>
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
