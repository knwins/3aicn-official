<!DOCtYPE html PUBliC "-//W3C//Dtd Xhtml 1.0 transitional//EN" "http://www.w3.org/tr/xhtml1/Dtd/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!--#include file="fconn.asp"-->
<%set rs=server.createobject("adodb.recordset")
sql="select * from site where id=1"
rs.open sql,conn,1,1
title=rs("title")
description=rs("description")
key=rs("key")
rs.close
set rs=nothing%>
<head>
<title><%=title%></title>
<meta http-equiv="Content-type" content="text/html; charset=gb2312" />
<meta name="description" content="<%=description%>">
<meta name="keywords" content="<%=key%>">
<link rel="stylesheet" type="text/css" href="css/style.css">
<script src="js/serviceqq.js" language="javascript" charset="gb2312" type='text/javascript'></script>
<script src="js/gonggao.js" language="javascript" type='text/javascript'></script>
<script src="js/ggao.js" language="javascript" type='text/javascript'></script>
</head>
<body>
<div id="container">
<div class="ggao" id=visible><b>��վ���棺</b><a id=hottext style="color:#FFFFFF"></a><div id=incoming style="DISPLAY: none">
  <div id=AllNews>
  <%
set rs=server.createobject("adodb.recordset")
sql="select top 10 * from news where ver='cn' and type=34 order by orders desc"
rs.open sql,conn,1,1				 	
 i=0
do while not rs.eof%>  
    <div id=1>
      <div id=Summary><%=rs("product_name")%> - <%=rs("createdate")%></div>
      <div id=NewsLink><%=rs("product_id")%>.html</div>
    </div>
<%rs.movenext
i=i+1
if i=8 then exit do
loop
rs.close
set rs=nothing %>
  </div>
  <div id=AddNews></div>
</div></div>
  <div id="header" style="height:140px;">
    <div class="logo"><img src="images/logo.gif"></div>
    <h1 style="position:absolute; top:74px; left:24px; color:#000; font-weight:normal; font-size:12px; margin:0; padding:0;"> <%=description%></h1>
    <div class="language" style="top:15px;">
      <p> <a href="/" target="_self">���İ�</a> <a href="/en">ENGLISH</a></p>
    </div>	
    <div class="tel">�ͷ���4000442113</div>
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
    <div id=imgADPlayer style="padding-bottom:5px;" align="center">
 <SCRIPT type="text/javascript">                     
/*
*��ȫ�������Ϊhtml��ֱ��ʹ�ã���jsp������
*/
var widths=950;    /*��ʾ�߶�*/                  
var heights=250;   /*��ʾ���*/                   
var counts=3;      /*��Ƭ����*/       


//һ��img��Ӧһ������url����������ȡ��ʱ���Ƕ�Ӧ�ģ��������Ķ�������ҲҪ��Ӧ�Ķ�
//img.src��ָͼƬ·����url�ǵ��ͼƬ����ת��ҳ������
<% set rs5=server.createobject("adodb.recordset")
rs5.open"select * from banner where ver='cn' order by ID asc",conn,1
j=1  
while not rs5.eof
 %>
		img<%=j%>=new Image ();img<%=j%>.src='<%=rs5("p_image")%>';
		url<%=j%>=new Image ();url<%=j%>.src='<%=rs5("url")%>';
		<%
		 rs5.movenext
		 j=j+1
wend
 
rs5.close
set rs5=nothing%>

/* ���»�������Ҫ�Ķ������������漰���Ķ������Ͳ�����ʱ�� ����ȻҲ�иĶ��ı�Ҫ*/
var nn = 1;
var key = 0;
function change_img() {
if (key == 0) {
   key = 1;
} else {
   if (document.all) {
    document.getElementById("pic").filters[0].Apply();
    //ͼƬ�л��ı��е�ʱ�䣬ԽС�л�Խ��
    document.getElementById("pic").filters[0].Play(duration = 2);
   }
}
eval("document.getElementById(\"pic\").src=img" + nn + ".src");
eval("document.getElementById(\"url\").href=url" + nn + ".src");
for (var i = 1; i <= counts; i++) {
document.getElementById("xxjdjj" + i).className = "axx";
}
document.getElementById("xxjdjj" + nn).className = "bxx";
nn++;
if (nn > counts) {
   nn = 1;
}
//ͼƬ�л���ʱ����
tt = setTimeout("change_img()", 10000);
}
function changeimg(n) {
nn = n;
window.clearInterval(tt);
change_img();
}

document.write("<style>");
document.write(".axx{padding:1px 7px 1px 7px;*padding:1px 7px 1px 7px;border:#fff 0px solid;}");
document.write("a.axx:link,a.axx:visited{text-decoration:none;color:#333;line-height:13px;font:12px sans-serif;background-color:#cccccc;margin-right:1px;margin-left:1px;filter:alpha(style=1,opacity=80,finishOpacity=80);-moz-opacity:0.8; opacity:1;}");
document.write("a.axx:active,a.axx:hover{text-decoration:none;color:#333;line-height:13px;font:12px sans-serif;background-color:#999;margin-right:1px;margin-left:1px;}");
document.write(".bxx{padding:1px 7px 1px 7px;;*padding:padding:1px 7px 1px 7px;;border:#fff 0px solid;}");
document.write("a.bxx:link,a.bxx:visited{text-decoration:none;color:#fff;line-height:13px;font:12px sans-serif;background-color:#999;margin-right:1px;margin-left:1px;}");
document.write("a.bxx:active,a.bxx:hover{text-decoration:none;color:#fff;line-height:13px;font:12px sans-serif;background-color:#999;margin-right:1px;margin-left:1px;}");
document.write("</style>");
document.write("<div style=\"width:" + widths + "px;height:" + heights + "px;overflow:hidden;text-overflow:clip;\">");
document.write("<div><a target='_blank' id=\"url\"><img id=\"pic\" style=\"border:0px;filter:progid:dximagetransform.microsoft.wipe(gradientsize=1.0,wipestyle=4, motion=forward)\" width=" + widths + " height=" + heights + " /></a></div>");
document.write("<div style=\"filter:alpha(style=1,opacity=0,finishOpacity=100);-moz-opacity:0.8; opacity:1;width:100%;text-align:right;top:-18px;position:relative;margin-top:0px;margin-bottom:0px;margin-right:2px;height:20px;padding:0px;border:0px;\">");

for (var i = 1; i < counts + 1; i++) {
document.write("<a href=\"javascript:changeimg(" + i + ");\" id=\"xxjdjj" + i + "\" class=\"axx\" target=\"_self\">" + i + "</a>");
}

document.write("</div></div>");

change_img();
</SCRIPT>
  </div>
  <div id="maincontent">
    <div class="boxes_1">
      <!--��Ʒ�б� begin-->
      <div class="box" >
            <div class="boxtitle"><a>����������Ʒ����</a></div>
            <div style="width:239px; height:240px; padding:6px 2px; border:1px #ff0000 solid; border-top:0px; margin-top:-1px;">
              <div class="indexprolist">
                <table width="93%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <%
	set rs=server.createobject("adodb.recordset")
	rs.open "select * from product_cate where root_id=0 and ver='cn' order by listorder asc",conn,1
	while not rs.eof
		getcateid = rs("pro_cate_id")
		getcatename = rs("pro_cate_name")
%>
                  <tr>
                    <td ><img src="images/dianxian.gif" /></td>
                  </tr>
                  <tr>
                    <td height="24" align="left" class="big"><img src="images/big.gif" /> <a href="probig_<%=rs("pro_cate_id")%>.html"> <%=rs("pro_cate_name")%></a></td>
                  </tr>
                  <tr>
                    <td><img src="images/dianxian.gif" /></td>
                  </tr>
                  <%
								set rs_pro=server.createobject("adodb.recordset")
								sql = "select * from product where ver='cn' and BIGTYPE = "& getcateid &" and TYPE = 133 order by PRODUCT_ID asc"
								rs_pro.open sql,conn,1
								if  not rs_pro.eof  then
									while not rs_pro.eof 
%>
                  <tr>
                    <td height="22" align="left" class="small"><img src="images/small.gif" width="15" height="15" /> <a href="pro_<%=rs_pro("PRODUCT_ID")%>.html"><%=rs_pro("PRODUCT_NAME")%></a> </td>
                  </tr>
                  <tr>
                    <td><img src="images/dianxian.gif" /></td>
                  </tr>
                  <%                          
										rs_pro.movenext
									wend 
								end if
								rs_pro.close
								set rs_pro=nothing
%>
                  <%
		rs.movenext
	wend 
	rs.close
	set rs=nothing
%>
                </table>
              </div>
              <div class="indexprolist"></div>
            </div>
      </div>
      <div class="box2" style="height:290px; right:0; top:0;" >
        <div class="boxtitle2"><a >���������Ƽ���Ʒ</a> <span><a href="product.html">����&gt;&gt;</a></span> </div>
        <div class="boxcont2" style="padding-top:12px;">
          <% set rs=server.createobject("adodb.recordset")
sqlnew="select * from product where ver='cn' and commend=1 order by orders" 
rs.open sqlnew,conn,1,1	
i=0				 	
do while not rs.eof %>
          <div style="text-align:center; width:164px; float:left; margin:4px 2px 2px 2px;*margin:4px 2px 1px 2px;"> <a href="pro_<%=rs("product_id")%>.html"><img src="<%=rs("s_image") %>" width="140" height="130" border="0"  style="border:1px silver solid; padding:2px;" /></a>
            <div style="text-align:center;"><a href="pro_<%=rs("product_id")%>.html"><b><%=left(rs("product_name"),10)%></b></a></div>
            <div style="font-size:12px;text-align:left; width:156px;"><%=left(rs("intro"),40)%></div>
          </div>
          <%                          
					rs.movenext
				i=i+1
if i=4 then exit do
loop
			rs.close
			set rs=nothing
%>
        </div>
      </div>
    </div>
    <div class="boxes">
      <div class="box">
        <div class="boxtitle"><a>�����������̰���</a> <span><a href="caselist.html">����&gt;&gt;</a></span></div>
        <div style="width:239px; height:300px; padding:6px 2px; border:1px #FF0000 solid; border-top:0px; margin-top:-1px;">
          <div style="text-align:center; margin:4px auto;">
            <div style="position:relative; left:0; top:0; width:220px; height:150px; background:#000">
              <div id="fv2_filbox">
                <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" id=scriptmain name=scriptmain codebase="http://download.macromedia.com/pub/shockwave/cabs/
flash/swflash.cab#version=6,0,29,0" width="230" height="150">
      <param name="movie" value="swf/bcastr.swf?bcastr_xml_url=xml.asp">
      <param name="quality" value="high">
      <param name=scale value=noscale>
      <param name="LOOP" value="false">
      <param name="menu" value="false">
      <param name="wmode" value="transparent">
      <embed src="swf/bcastr.swf?bcastr_xml_url=xml.asp" width="230" height="150" loop="false" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" salign="t" name="scriptmain" menu="false" wmode="transparent"></embed>
      <param name="wmode" value="transparent" />
    </object>
              </div>
            </div>
          </div>
          <div class="indexprolist" style="margin-top:4px;">
            <ul>
             <%
set rs=server.createobject("adodb.recordset")
sql="select * from news where ver='cn' and type=18 order by orders desc"
rs.open sql,conn,1,1
i=0				 	
do while not rs.eof%>
              <li style="line-height:24px;"><a href="<%=rs("product_id")%>.html" target="_blank"> <%=left(rs("product_name"),16)%></a></li>
              <%rs.movenext
i=i+1
if i=6 then exit do
loop
rs.close
set rs=nothing %>
            </ul>
          </div>
        </div>
      </div>
      <div class="box2"  style="width:450px; height:350px; right:247px; top:0; ">
        <div class="boxtitle2"><a >����������ƷӦ��</a> <span><a href="solutions.html">����&gt;&gt;</a></span> </div>
        <div class="boxcont2">
         
		 <% set rs=server.createobject("adodb.recordset")
sql="select top 6 * from news where ver='cn' and url<>'' and type=33 order by hot desc"
rs.open sql,conn,1,1	
i=0				 	
do while not rs.eof %>
          <div style="text-align:center; width:134px; float:left; margin:4px 2px 2px 2px;*margin:4px 2px 1px 2px;"> <a href="<%=rs("product_id")%>.html"><img src="<%=rs("url") %>" width="119" height="109" border="0"  style="border:1px silver solid; padding:2px;" /></a>
            <div style="text-align:center;"><a href="<%=rs("product_id")%>.html"><%=left(rs("product_name"),10)%></a></div>
          </div>
          <%                          
					rs.movenext
				i=i+1
if i=6 then exit do
loop
			rs.close
			set rs=nothing
%>          
        </div>
      </div>
      <div class="box2" style="width:240px; height:350px; right:0; top:0;">
        <div class="boxtitle2"><a >����������ҵ��̬</a> <span><a href="news.html">����&gt;&gt;</a></span> </div>
        <div class="boxcont2">
          <div style="padding:0px 0px  6px 0px;">
            <%
set rs=server.createobject("adodb.recordset")
sql="select top 1 * from news where ver='cn' and type=20 order by orders desc"
rs.open sql,conn,1,1				 	
 i=0
do while not rs.eof%>
            <h2><a href="<%=rs("product_id")%>.html" target="_blank"> <%=left(rs("product_name"),32)%></a></h2>
            <p><%=left(rs("intro"),80)%> </p>
            <%rs.movenext
i=i+1
if i=2 then exit do
loop
rs.close
set rs=nothing %>
          </div>
          <DIV id=Marquee_3_1 style="OVERFLOW: hidden; WIDTH:230px; HEIGHT:200px;">
            <DIV id=Marquee_3_2 align=left>
              <div style="width:230px; height:auto; ">
                <%
set rs=server.createobject("adodb.recordset")
sql="select top 10 * from news where ver='cn' and type=20 order by orders asc"
rs.open sql,conn,1,1				 	
 i=0
do while not rs.eof%>
                <p style="text-align:left;font-size:12px;">&raquo;<a href="<%=rs("product_id")%>.html" target="_blank" class="mostread"> <%=left(rs("product_name"),16)%></a></p>
                <%rs.movenext
i=i+1
if i=8 then exit do
loop
rs.close
set rs=nothing %>
              </div>
            </DIV>
            <DIV id=Marquee_3_3></DIV>
          </DIV>
        </div>
      </div>
      <!--չ����Ϣ  end-->
    </div>
    <div style="margin:0px 0px 8px 0px;"><img src="images/pic.jpg" width="950" height="150" /></div>
    <!--3 begin-->
    <div class="boxes">
      <!--��ϵ���������Ƽ� begin-->
      <div class="box">
        <div class="boxtitle"><a>����������ϵ��ʽ</a> <span><a href="contactus.html">����&gt;&gt;</a></span> </div>
        <div class="boxcont" style="padding-top:9px;">
          <p><b>�㶫���ڹ�˾</b></p>
          <p>��ַ�������и����������빤ҵ<br>
            ��526��618<br />
            �ͷ���4000442113
<br />
            ���棺+86-755-28260069
<br />
            ���ʣ�goumike@163.com<br />
          �ͷ�24Сʱ���ߣ�13600442113</p>
          <p><b>��ַ��<a href="Http://www.3aicn.com">Http://www.3aicn.com</a></b></p>
          <p><img src="images/contactus.gif" width="212" height="123"></p>
        </div>
      </div>
      <div class="box2"  style="width:450px; height:350px; right:247px; top:0; ">
        <div class="boxtitle2"><a >����������˾���</a> <span><a href="aboutus.html">����&gt;&gt;</a></span> </div>
        <div class="boxcont2">
          <p><img src="images/2314631833.jpg" alt="" width="413" height="123" border="0" /></p>
          <p> ���������������������޹�˾��רҵ�ķ����¡��ܷ�ϵͳ�Ĺ��̹�˾��ͬʱ�ṩ���ֲ�Ʒ���ͻ�ѡ��,Ϊ�ͻ��ṩ���Ƶ�ʩ����ѵ���ۺ����ͼ���֧�֡���Ҫ��Ʒ�У�UL��֤��FM��֤�ĵ��Է����ܷ⽺��UL��֤��FM��֤�ķ����ܷ⽺��UL��֤��FM��֤�ķ������ͷ����ࡢUL��֤��FM��֤�������ͷ����ࡢUL��֤��FM��֤������������ִ������Ȧ��UL��֤��FM��֤����ĭ��²��ϡ�UL��֤��FM��֤�ķ��𸴺ϰ塢UL��֤��FM��֤�İ��������ġ�UL��֤��FM��֤�ĵ��·���Ϳ�ϡ�UL��֤��FM��֤�ĸֽṹ����Ϳ�ϡ�</p>
          <p> �� Ϊ ר ҵ �� �� �� �� �� �� �� !</p>
        </div>
      </div>
      <div class="box2" style="width:240px; height:350px; right:0; top:0;">
        <div class="boxtitle2"><a>���������������</a></div>
        <div class="boxcont2">
          <DIV id=Marquee_4_1 style="OVERFLOW: hidden; WIDTH:230px;">
          <p align="center"><a href="http://www.baidu.com" target="_blank"><img src="images/logo/baidu.gif" width="190" height="90" border="0"></a><br>
          </p><p align="center"><a href="http://www.soso.com" target="_blank"><img src="images/logo/soso.gif" width="190" height="60" border="0"></a><br>
            <br>
          </p>
            <p align="center"><a href="http://www.google.com.hk" target="_blank"><img src="images/logo/google.gif" width="190" height="67" border="0"></a></p>
          </DIV>
        </div>
      </div>
    </div>
    <!--3 end-->
<div id="footer">
      <div class="footernav"> <a href="/index.html">��ҳ</a>| <a href="probig_78.html">��Ʒ����</a>| <a href="caselist.html">���̰���</a>| <a href="rongyu.html">��ҵ����</a>| <a href="qanda.html">��������</a>| <a href="aboutus.html">��������</a>| <a href="contactus.html">��ϵ��ʽ</a> </div>
      <p style=" padding-top:5px;">Copyright &copy; 2003-2011 ���������������������޹�˾, All Rights Reserved &nbsp;&nbsp; ��ICP��06023278�� <br />
        �㶫���ڣ� �����и����������빤ҵ��526��618  �ͷ���4000442113  ���棺0755-28260069  �ֻ���13600442113 <img width="0" height="0" src="create_index.asp?htime=<%=now%>"><br>
<%=description%>
<script src="http://s16.cnzz.com/stat.php?id=3224905&web_id=3224905&show=pic" language="JavaScript"></script> </p>
    </div>
  </div>
</div>
</body>
</html>
