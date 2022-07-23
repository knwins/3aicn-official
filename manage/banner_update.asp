<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<% 
	dim rs, id, title, p_image, url
	id  = replace(request("id"),"'","")
	set rs=server.createobject("adodb.recordset")
	rs.open "select * from banner where id = "&id&"",conn,1
	title = rs("title")
	p_image      = rs("p_image")
	url       = rs("url")
	rs.close
	set rs = nothing
%>
<html>
<head>
<title>后台管理系统</title>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
</head>
<body bgcolor="#ffffff" leftmargin=0 topmargin=0>
<br>
<table border="1" width="90%" align="center" bordercolorlight="#c0c0c0" cellspacing="1" cellpadding="3" bordercolordark="#ffffff" >
  <form name="myform" action=banner_upd.asp method=post>
    <tr>
      <td width="20%" align="center">标题：</td>
      <td width="80%"><input name=title id="name" value="<%=title%>" size="60"></td>
    </tr>
    <tr align="center">
      <td>图片信息：</td>
      <td align="left" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><input name=p_image value="<%=p_image%>" size=35 maxlength=50>
          <input name="Submit" type="button" onClick="window.open('../upload_flash.asp?formname=myform&editname=image&uppath=pro_img/banner&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')" value="上传">
          <br>
          <br>
      <font color=red>（注：不支持中文名且小于200KB 图片大小：970*240 ）</font> </td>
    </tr>
    <tr align="center">
      <td>图片链接：</td>
      <td width="80%" align="left" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><input name=url id="url" value="<%=url%>" size="60"></td>
    </tr>
    <tr>
      <td class="font" height="28" width="80%" colspan="2" ><p align="center">
         <input name=id value="<%=id%>" type="hidden">
		 <input type="submit" name="add" value=" 添加 " >
          <input type="reset" name="reset" value=" 重写 ">
      </td>
    </tr>
  </form>
</table>
</body>       
</html>

