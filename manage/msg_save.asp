<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
id=request("id")
userid=request("userid")
page=request("page")
set rs=server.CreateObject("adodb.recordset")
sql="select * from msg where id="&id
rs.open sql,conn,3,2 
	rs("sendmsg")=trim(request("sendmsg")) '��Ʒ����
	rs("senddate")=date()
rs.update
%>
<SCRIPT LANGUAGE="JavaScript">

  window.alert("�ѳɹ��޸ģ�");
  window.location.href="msg.asp?page=<%=page%>";

</SCRIPT>