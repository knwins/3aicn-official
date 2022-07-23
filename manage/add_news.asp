<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
</head>
<body>
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="28"><font color="#FF0000">当前位置 --＞增加企业信息</font>　　　　（带“<font color="#FF0000">*</font>”为必填项） </td>
  </tr>
  <tr>
    <td height="1" bgcolor="#4E6A82"></td>
  </tr>
</table>
<table border="1" width="90%" align="center" bordercolorlight="#c0c0c0" cellspacing="1" cellpadding="3" bordercolordark="#ffffff" >
  <form name="myform" action=add_news_ok.asp method=post>
    <tr>
      <td width="20%">序号：</td>
      <td width="80%"><input name=orders size="40" value="0">
        <font color=red>（必须为数字，数字越小越靠前）</font></td>
    </tr>
    <tr>
      <td width="20%" align="left">新闻标题：</td>
      <td width="80%"><input name=product_name id="product_name" size="60"></td>
    </tr>
    <tr>
      <td align="left">新闻图片：</td>
      <td width="562" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><input name=s_image size=35 maxlength=50 value="">
          <input name="Submit" type="button" onClick="window.open('../upload_flash.asp?formname=myform&editname=s_image&uppath=pro_img/s&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')" value="小图">
          <font color=red>（注：不支持中文名且小于200KB）</font> </td>
    </tr>
    <tr>
      <td width="20%" align="left">类别：</td>
      <td width="80%"><select name="types">
          <%call showtype()%>
        </select>      </td>
    </tr>
    <tr align="center">
      <td align="left">新闻说明：</td>
      <td align=left valign="top" class="unnamed2" bgcolor="#ffffff"><textarea name="intro" cols="50" rows="4" id="intro"></textarea></td>
    </tr>
    <tr align="center">
      <td>内容：</td>
       <td align=left valign="top" class="unnamed2" bgcolor="#ffffff"><input type="hidden" name="content" id="content"><iframe id="myiframe" src="edit/ewebeditor.htm?id=content&ReadCookie=0" frameBorder="0" marginHeight="0" marginWidth="0" scrolling="No" width="621" height="459"></iframe></td>
    </tr>
    <tr>
      <td class="font" height="28" width="20%" colspan="2" ><p align="center">
          <input type="submit" name="add" value=" 添加 " >
          <input type="reset" name="reset" value=" 重写 ">
      </td>
    </tr>
  </form>
</table>
</body>
</html>
<%
sub showtype()
	set rscate=server.createobject("adodb.recordset")
	rscate.open "select * from news_cate where ver='"&session("ver")&"' order by listorder asc",conn,1
	while not rscate.eof
		str1="<option value="""& rscate("id") & """>" & rscate("typename") & "</option>"
		response.write str1
		rscate.movenext
	wend 
	rscate.close
	set rscate=nothing
end sub
%>
