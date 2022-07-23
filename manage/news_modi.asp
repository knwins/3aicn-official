<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<% 
	dim rs, product_id, product_name, content, intro, commend, orders
	product_id  = replace(request("id"),"'","")
	page=clng(request("topage"))
	typeids=request("typeids")

	set rs=server.createobject("adodb.recordset")
	rs.open "select * from news where product_id = "& product_id &" order by orders,createdate desc",conn,1
	product_name = rs("product_name")
	content      = rs("content")
	orders       = rs("orders")
	types        = rs("type")
	intro        = rs("intro")
	s_image        = rs("url")
	createdate   = rs("createdate")
	rs.close
	set rs = nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="35"><font color="#FF0000">当前位置：<%=request("nbmc")%> --＞企业信息管理　　</font>（带“<font color="#FF0000">*</font>”为必填项） </td>
  </tr>
</table>
<table border="1" width="90%" align="center" bordercolorlight="#c0c0c0" cellspacing="1" cellpadding="3" bordercolordark="#ffffff" >
  <form name="myform" action="news_modi_ok.asp" method=post>
    <input type=hidden name="product_id" value="<%=product_id%>">
	<input type=hidden name="typeids" value="<%=typeids%>">
	<input type=hidden name="page" value="<%=page%>">
    <tr>
      <td>序号：</td>
      <td><input name=orders size="40" value="<%=orders%>">
          <font color=red>（必须为数字，数字越小越靠前）</font></td>
    </tr>
    <tr>
      <td width="12%">新闻标题：</td>
      <td width="88%"><input name=product_name size="60" value="<%=product_name%>"></td>
    </tr>
    <tr>
      <td align="left">新闻图片：</td>
      <td bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"> <input name="s_image" size=35 maxlength=50 value="<%=s_image%>">
              <input name="Submit" type="button" onClick="window.open('../upload_flash.asp?formname=myform&editname=s_image&uppath=pro_img/news&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')" value="小图">
          <font color=red>（注：不支持中文名且小于200KB）</font> </td>
    </tr>
    <tr>
      <td width="12%">类别：</td>
      <td width="88%"><select name="types">
        <%call getchiled(cint(types),1)%>
      </select>      </td>
    </tr>
    <tr align="center">
      <td align="left">新闻说明：</td>
      <td align=left valign="top" class="unnamed2" bgcolor="#ffffff"><textarea name="intro" cols="50" rows="4" id="intro"><%=intro%></textarea></td>
    </tr>
    <tr align="center">
      <td>内容：</td>
      <td align=left valign="top" class="unnamed2" bgcolor="#ffffff"><textarea name="content" id="content" style="display:'none'"><%=content%></textarea><iframe id="myiframe" src="edit/ewebeditor.htm?id=content&ReadCookie=0" frameBorder="0" marginHeight="0" marginWidth="0" scrolling="No" width="621" height="459"></iframe>      </td>
    </tr>
    <tr>
      <td class="font" height="28" colspan="2" >
	  <p align="center">
        <input type="submit" name="add" value=" 修改 " >
        <input type="button" name="back" value=" 返回 " onClick="history.back();">
      </td>
    </tr>
  </form>
</table>
</body>
</html>

<%
sub	getchiled(cateid,level)
		set rschiled=server.createobject("adodb.recordset")
		rschiled.open "select * from news_cate where ver='"&session("ver")&"' order by listorder desc",conn,1
		while not rschiled.eof
			if rschiled("id") = cateid then
				str1="<option value="""& rschiled("id") &  """selected>" 
			else
				str1="<option value="""& rschiled("id") &  """>" 
			end if
			call list(rschiled("typename"), str1,level)
			rschiled.movenext
		wend
		rschiled.close
		set rschiled=nothing
end sub

sub	list(catename, str,level)
		dim space
		for i=1 to level
			space=space+""
		next
		str2=catename & "</option>"
		str=str & space & str2
		response.write str
end sub
%>
