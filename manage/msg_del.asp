<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
	id=request("id")
	page=request("page")
	userid=request("userid")
	dim sql  
	sql = "delete * from [msg] where id = "&id
	conn.Execute (sql)

%>
<SCRIPT LANGUAGE="JavaScript">

  window.alert("�Ѿ�ɾ����");
  window.location.href="msg.asp?page=<%=page%>";

</SCRIPT>

<%CloseDatabase%>