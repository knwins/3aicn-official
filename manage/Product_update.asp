<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<% 
	dim rs, product_id, PRODUCT_NAME, CONTENT, INTRO, COMMEND, ORDERS
	session("page")=request("page")
	product_id  = replace(Request("id"),"'","")
	set rs=server.CreateObject("ADODB.RecordSet")
	rs.open "select * from product where product_id = "& product_id &" order by orders,createdate desc",conn,1
	typenames = rs("typename")
	PRODUCT_NAME = rs("PRODUCT_NAME")
	CONTENT      = rs("CONTENT")
	ORDERS       = rs("ORDERS")
	INTRO        = rs("INTRO")
	COMMEND      = rs("COMMEND")
	S_IMAGE      = rs("S_IMAGE")
	B_IMAGE      = rs("B_IMAGE")
	TYPES        = rs("TYPE")
	id           = rs("PRODUCT_ID")
	rs.close
	set rs = nothing
%>
<html>
<head>
<title>��̨����ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet type="text/css" href="js/media.css">
<link rel=stylesheet type="text/css" href="js/form_media.css">
<SCRIPT LANGUAGE=JAVASCRIPT>
<!--
	function checkform() {
		if (document.myform.PRODUCT_NAME.value==""){
			window.alert("����Ʒ���ơ�����Ϊ�գ�");
			document.myform.PRODUCT_NAME.onfocus;
			return (false);
		}
		if (document.myform.ORDERS.value==""){
			window.alert("����š�����Ϊ�գ�");
			document.myform.ORDERS.onfocus;
			return (false);
		}
		
		return true
	}
//-->
</SCRIPT>
</head>
<body bgcolor="#FFFFFF" >
<script language="JavaScript"> 
function ConfirmDel() 
{ 
if (confirm("�����Ҫɾ�����ļ���!")) 
return true; 
else 
return false; 
} 
</script>
<table border="1" width="90%" align="center" bordercolorlight="#C0C0C0" cellspacing="1" cellpadding="3" bordercolordark="#FFFFFF" >
              <form name="myform" action=Product_upd.asp method=post onSubmit="return checkform();">
              <input type=hidden name="product_id" value="<%=product_id%>">
                <tr>
                  <td>��ţ�</td>
                  <td>
                  	<input name=ORDERS size="40" value="<%=ORDERS%>" onKeyUp="value=value.replace(/\D+/g,'')">
                  	<font color=red>������Ϊ���֣�����ԽСԽ��ǰ��</font>                  </td>
				</tr>
								 <tr>
                  <td>��Ʒ��������</td>
                  <td><input name=typename id="typename" size="20" value="<%=typenames%>">
                    <font color=red>��Ҫ�ڲ�Ʒ��������ʾ12��Ӣ���ַ���</font></td>
                </tr>
                <tr>
                  <td width="20%">��Ʒ���ƣ�</td>
                  <td width="90%"><input name=PRODUCT_NAME size="40" value="<%=PRODUCT_NAME%>"></td>
				</tr>
          <tr>
                  <td width="20%">��ƷСͼƬ��</td>
            <td>
              <input name=s_image size=35 maxlength=50 value="<%=S_IMAGE%>">
              <input name="Submit" type="button" onClick="window.open('../upload_flash.asp?formname=myform&editname=s_image&uppath=pro_img/s&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')" value="Сͼ">
            <font color=red>��ע��СͼƬ��X��Ϊ200X100����,��֧����������</font>                  </td>
  				</tr>
				                 <tr>
                  <td width="20%">��Ʒ��ͼƬ��</td>
                  <td width="80%">
                  <input name=B_IMAGE size=35 maxlength=50 value="<%=B_IMAGE%>">
                  <input name="Submit" type="button" onClick="window.open('../upload_flash.asp?formname=myform&editname=B_IMAGE&uppath=pro_img/b&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')" value="��ͼ">
                  <font color=red>��ע����ͼƬ��Ϊ500���أ����Ȳ���,��֧����������</font></td> 
				</tr>
                <tr>
                  <td width="20%">���</td>
                  <td width="90%">
                  <SELECT name="types"><%call showtype(cint(types))%></SELECT>                  </td>
				</tr>
                <tr>
      <td>��Ҫ˵����</td>
      <td><textarea name="INTRO" cols="45" rows="5" wrap="physical" id="INTRO"><%=INTRO%></textarea></td>
    </tr>
                <tr>
                  <td width="20%">�Ƿ��Ƽ���Ʒ��</td>
                  <td width="90%"><%if COMMEND=1 then %><input name=COMMEND type="checkbox" value="checkbox" checked><%else%>                    <input name=COMMEND type="checkbox" value="checkbox"><%end if%>
                  �������Ƿ����Ƽ���Ʒ����ʾ��</td>
				</tr>
                <tr align="center"> 
                  <td>����˵����</td>
                  <td align=left valign="top" class="unnamed2" bgcolor="#FFFFFF"> 
                    <textarea name="content" id="content" style="display:'none'"><%=content%></textarea><iframe id="myiframe" src="edit/ewebeditor.htm?id=content&ReadCookie=0" frameBorder="0" marginHeight="0" marginWidth="0" scrolling="No" width="621" height="459"></iframe>                  </td>
                </tr>
                 
                <tr>              
                  <td class="font" height="28" width="20%" colspan="2" >    
                    <p align="center">   
                    <input type="submit" name="add" value=" �޸� " >
                    <input type="button" name="back" value=" ���� " onClick="history.back();">
                  </td>
                </tr>
              </form>
</table>       
</body>       
</html>
<%
sub showtype(cateid)	
	set rscate=server.CreateObject("ADODB.RecordSet")
	rscate.open "select * from product_cate WHERE ver='"&session("ver")&"' and ROOT_ID=0 order by listorder ASC",conn,1
 
	if rscate("Pro_Cate_ID") = cateid then
	str3="<option value="""& rscate("Pro_Cate_ID") &  """selected>" & rscate("Pro_Cate_Name") & "</option>"
	else
	set rscates=server.CreateObject("ADODB.RecordSet")
	rscates.open "select * from product_cate where ver='"&session("ver")&"' and pro_cate_id="&cateid&" order by listorder ASC",conn,1
	str3="<option value="""&rscates("pro_cate_id")&""" selected>&nbsp;&nbsp;" & rscates("Pro_Cate_Name") & "</option>"
	rscates.close
	set rscates=nothing
	end if
	response.Write str3
	
	while not rscate.eof	
	str1="<optgroup label=" & rscate("Pro_Cate_Name") & "></optgroup>"
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
			if rschiled("pro_cate_id") = cateid then
				str1="<option value="""& rschiled("pro_cate_id") &  """selected>" 
			else
				str1="<option value="""& rschiled("pro_cate_id") &  """>" 
			end if
			call list(rschiled("pro_cate_name"), str1,level)
			call Getchiled(rschiled("pro_cate_id"),(level+1))
			rschiled.movenext
		wend
		set rschiled=nothing
end sub
sub	list(catename, str,level)
		dim space
		for i=1 to level
			space=space+"��"
		next
		str2=catename & "</option>"
		str=str & space & str2
		Response.Write str
end sub
%>
<%CloseDatabase%>