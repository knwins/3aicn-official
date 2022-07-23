<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<html>
<head>
<title>后台数据管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<SCRIPT LANGUAGE=JAVASCRIPT>
<!--
function checkform() {
	if (document.myform.PRODUCT_NAME.value==""){
		window.alert("“产品名称”不能为空！");
		document.myform.PRODUCT_NAME.onfocus;
		return (false);
	}
	if (document.myform.ORDERS.value==""){
		window.alert("“编号”不能为空！");
		document.myform.ORDERS.onfocus;
		return (false);
	}
	return(true);
}
//-->
</SCRIPT>

</head>
<body bgcolor="#FFFFFF" leftmargin=0>
<table width="90%" border="1" align="center" cellpadding="3" cellspacing="1" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" >
              <form name="myform" action=Product_add.asp method=post onSubmit="return checkform();">
                <tr>
                  <td width="161">编号：</td>
                  <td width="562">
                  	<input name=ORDERS size="40" value="0" onKeyUp="value=value.replace(/\D+/g,'')">
                  	<font color=red>（必须为数字，数字越小越靠前）</font></td>
				</tr>
                <tr>
                  <td>产品类型名：</td>
                  <td><input name=typename id="typename" size="20">
                    <font color=red>主要在产品分类里显示12个英文字符内</font></td>
                </tr>
                <tr>                   
                  <td width="161">产品名称：</td>
                  <td width="562"><input name=PRODUCT_NAME size="40"></td>
				</tr>
                <tr>
                  <td width="161">产品小图片：</td>
                  <td width="562">
                  <input name=s_image size=35 maxlength=50 value="">
                  <input name="Submit" type="button" onClick="window.open('../upload_flash.asp?formname=myform&editname=s_image&uppath=pro_img/s&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')" value="小图">
                  <font color=red>（注：不支持中文名且小于200KB）</font>                  </td>
  				</tr>    <tr>
                  <td width="20%">产品大图片：</td>
                  <td width="80%">
                  <input name=B_IMAGE size=35 maxlength=50 value="">
                  <input name="Submit" type="button" onClick="window.open('../upload_flash.asp?formname=myform&editname=B_IMAGE&uppath=pro_img/b&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')" value="大图">
                  <font color=red>（注：大图片宽为500像素，长度不限,不支持中文名）</font></td> 
				</tr>
                <tr>
                  <td width="161">类别：</td>
                  <td width="562">
                  <SELECT name="TYPES"><%call showtype()%></SELECT>                  </td>
				</tr>
  							

                <tr>
                  <td width="161">是否推荐产品：</td>
                  <td width="562"><input type="checkbox" name=COMMEND value="checkbox">                    
                  （决定是否在推荐产品中显示）</td>
				</tr>
  							
								<tr align="center"> 
                  <td>内容：</td>
                  <td width="562" align=left valign="top" class="unnamed2"> 
                    <input type="hidden" name="content" id="content"><iframe id="myiframe" src="edit/ewebeditor.htm?id=content&ReadCookie=0" frameBorder="0" marginHeight="0" marginWidth="0" scrolling="No" width="621" height="459"></iframe></td>
                </tr>
                
                <tr>              
                  <td height="28" colspan="2" class="font" >    
                    <p align="center">   
                    <input type="submit" name="add" value=" 添加 " >
                    <input type="reset" name="reset" value=" 重写 ">
                  </td>     
                </tr>            
              </form>             
</table>       
<span class="unnamed2">
<input type="hidden" name="content">
</span>
</body>       
</html>       
<%
sub showtype()
	set rscate=server.CreateObject("ADODB.RecordSet")
	rscate.open "select * from product_cate WHERE ver='"&session("ver")&"' and ROOT_ID=0 order by listorder ASC",conn,1
	while not rscate.eof
		str1="<optgroup label="""& rscate("Pro_Cate_Name") & """></optgroup>"
		Response.Write str1
		call Getchiled(rscate("pro_cate_id"),1) 
		rscate.movenext
	wend 
	set rscate=nothing
end sub
sub	GetChiled(cateid,level)
		set rschiled=server.CreateObject("ADODB.RecordSet")
		rschiled.open "select * from product_cate WHERE ver='"&session("ver")&"' and ROOT_ID="&cateid&" order by listorder ASC",conn,1
		while not rschiled.eof
			if rschiled("pro_cate_id") =search2 then
				str1="<option value="""& rschiled("pro_cate_id") &  """selected>" 
			else
				str1="<option value="""& rschiled("pro_cate_id") &  """>" 
			end if
			call list(rschiled("Pro_Cate_Name"), str1,level)
			call Getchiled(rschiled("pro_cate_id"),(level+1))
			rschiled.movenext
		wend
		set rschiled=nothing
end sub
sub	list(catename, str,level)
		dim space
		for i=1 to level
			space=space+""
		next
		str2=catename & "</option>"
		str=str & space & str2
		Response.Write str
end sub
%>