<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!--#include file="fconn.asp"-->
<%

bigtype=request.QueryString("bigtype")
cateid=request.QueryString("cateid")
			
cname="产品中心"

IF 	cateid<>"" THEN
set RS_SMALL=server.CreateObject("ADODB.RecordSet")
RS_SMALL.open "select * from product_cate WHERE pro_cate_id= "&cateid &"" ,conn,1
PRODUCT_NAME_SMALL=RS_SMALL("PRO_CATE_NAME")
ROOT_ID=RS_SMALL("ROOT_ID")
RS_SMALL.close
set RS_SMALL=nothing

set RS_BIG=server.CreateObject("ADODB.RecordSet")
RS_BIG.open "select * from product_cate WHERE pro_cate_id= "&ROOT_ID &"" ,conn,1
PRODUCT_NAME_BIG=RS_BIG("PRO_CATE_NAME")
RS_BIG.close
set RS_BIG=nothing

END IF

IF 	bigtype<>"" THEN
set RS_BIG=server.CreateObject("ADODB.RecordSet")
RS_BIG.open "select * from product_cate WHERE pro_cate_id= "&bigtype &"" ,conn,1
PRODUCT_NAME_BIG=RS_BIG("PRO_CATE_NAME")
RS_BIG.close
set RS_BIG=nothing
END IF
	
set rs=server.createobject("adodb.recordset")
sql="select * from site where id=1"
rs.open sql,conn,1,1
title=rs("title")
description=rs("description")
key=rs("key")
rs.close
set rs=nothing%>
<head>
<title><%=title%>_产品展示</title>
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
      <p> <a href="/" target="_self">中文版</a> <a href="/en">ENGLISH</a></p>
    </div>	<div class="tel">客服：4000442113</div>
    <div id="navh2" style="top:100px;">
     <ul><li style="padding-left:10px;"><a href="index.html">首页</a></li>
        <li><a href="pro_616.html">3M防火板</a></li>
        <li><a href="pro_611.html">3M防火泥</a></li>
		 <li><a href="pro_610.html">3M防火密封胶</a></li>
        <li><a href="pro_612.html">3M阻火带</a></li>
        <li><a href="pro_615.html">3M电缆防火涂料</a></li>
        <li><a href="applications.html">3M防火封堵方案</a></li>
        <li><a href="gongyi.html">3M防火封工艺</a></li>
		<li><a href="caselist.html">3M防火封案例</a></li>
      </ul>
    </div>
  </div>
  <div style="margin:0px 0px 8px 0px;"> <img src="images/pic.jpg" /></div>
  <div id="maincontent">
    <table width="950" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div class="boxe_com">
            <div class="box" >
            <div class="boxtitle"><a>三爱电力产品分类</a></div>
            <div style="width:239px;padding:6px 2px; border:1px #ff0000 solid; border-top:0px; margin-top:-1px;">
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
                    <td height="26" align="left" class="big"><img src="images/big.gif" /> <a href="probig_<%=rs("pro_cate_id")%>.html"> <%=rs("pro_cate_name")%></a></td>
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
                    <td height="26" align="left" class="small"><img src="images/small.gif" width="15" height="15" /> <a href="pro_<%=rs_pro("PRODUCT_ID")%>.html"><%=rs_pro("PRODUCT_NAME")%></a> </td>
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
            <div class="box2_com">
              <div class="boxtitle3"><a ><b><%=cname %></b></a><span><a href="/" class="pathway">首页</a> >> 产品展示 >> <%=PRODUCT_NAME_BIG%> >> <%=PRODUCT_NAME_SMALL%></span></div>
              <div class="boxcont3" style="padding-top:12px;">
                <table cellspacing=0 cellpadding=0 width="100%" align=center border=0>
                  <tr>
                    <td align="center"><table cellspacing=1 cellpadding=2 width=100% border=0>
                        <%

set rs=server.CreateObject("ADODB.RecordSet")
if cateid<>"" then
sql = "select * from product where TYPE = "&cateid&" and ver='cn' order by orders asc"
elseif bigtype<>"" then
sql = "select * from product where bigtype = "& bigtype &" and ver='cn' order by orders asc"
else
sql = "select * from product where ver='cn' order by orders asc"
end if
rs.open sql,conn,1

if rs.eof and rs.bof then
response.write("<br>暂时没有记录<br>")
else
dim page
 IF not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Then
	page=1
	Else
	Page=Int(Abs(Request("page")))
 End if
rs.pagesize =12
total  = rs.RecordCount
mypagesize=rs.pagesize
rs.absolutepage = page	
%>
                        <tr>
                          <%
		Dim hang, idx
		hang = 3
		idx = 0
		while not rs.eof and mypagesize>0
			if rs("INTRO")="" then
			jiage = 0
			else
			jiage = rs("INTRO")
			end if
		%>
                          <td align="center" style="padding-top:15px;" valign="bottom"><table border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td bgcolor="#FFFFFF"><a href="pro_<%=rs("product_id")%>.html" target="_blank"> <img src="<%=rs("S_IMAGE")%>" width="181" height="174" border=0 style="border:1px silver solid; padding:2px;"></a> </td>
                              </tr>
                              <tr>
                                <td align="center" bgcolor="#FFFFFF" class="font"><a href="pro_<%=rs("product_id")%>.html" target="_blank"><b><%=rs("product_name")%></b></a></td>
                              </tr>
                            </table></td>
                          <%  
	mypagesize=mypagesize-1                        
	idx = idx + 1
		if idx mod hang = 0 then
		response.write "</tr><tr>"
		end if
	rs.MoveNext
	wend 
	%>
                        </tr>
                      </table></td>
                  </tr>
                  <tr>
                    <td height="52" align="center"><hr  style="color:#ff0000; height:1px;"/>
 <table border="0" cellspacing="0" cellpadding="0" class="font">
                        <tr>
                          <td align="center" height="22"><div id="pageClass3">
                              <ul>
                                <% for k=1 to rs.pagecount%>
                                <li class="currentpage"><a href="probig_<%=bigtype%>_<%=k%>.html"><%=k%></a></li>
                                <% next %>
                              </ul>
                            </div></td>
                        </tr>
                        <%
end if
rs.close
set rs=nothing%>
                      </table></td>
                  </tr>
                </table>
              </div>
            </div>
          </div></td>
      </tr>
      <tr>
        <td><div id="footer">
      <div class="footernav"> <a href="/index.html">首页</a>| <a href="product.html">产品中心</a>| <a href="caselist.html">工程案例</a>| <a href="rongyu.html">企业荣誉</a>| <a href="qanda.html">常见问题</a>| <a href="aboutus.html">关于我们</a>| <a href="contactus.html">联系方式</a> </div>
      <p style=" padding-top:5px;">Copyright &copy; 2003-2011 深圳市三爱电力技术有限公司, All Rights Reserved &nbsp;&nbsp; 粤ICP备06023278　 <br />
        广东深圳： 深圳市福田区八卦岭工业区526栋618  客服：4000442113
 传真：0755-28260069
 手机：13600442113<br />
      <%=description%></p>
    </div></td>
      </tr>
    </table>
  </div>
</div>
</body>
</html>
