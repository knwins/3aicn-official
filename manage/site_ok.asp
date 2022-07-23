<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
'----------------------------------------------------------------
' 获取数字型参数
'----------------------------------------------------------------
Function ReqNum ( StrName )
    ReqNum = Request ( StrName )
    if not isNumeric ( ReqNum ) then
        response.write "参数必须为数字型!"
        response.end
    end if
End Function

'----------------------------------------------------------------
' 获取字符型参数
'----------------------------------------------------------------
Function ReqStr ( StrName )
    ReqStr = Replace ( Request(StrName), "'", "''" )
End Function

session("id")=request.form("id")
set rs=server.CreateObject("adodb.recordset")
sql="select * from site where id="&session("id")
rs.open sql,conn,3,2
	rs("title")=request.form("title")
	rs("key")=request.form("key")
	rs("description")=request.form("description")
	rs("url")=request.form("url")
	rs("site")=request.form("site")
	rs("add")=request.form("add")
	rs("tel")=request.form("tel")
	rs("support")=request.form("support")
	rs("companytel")=request.form("companytel")
	rs("fox")=request.form("fox")
	rs("email")=request.form("email")
rs.update
%>
<SCRIPT LANGUAGE="JavaScript">

  window.alert("已成功修改该记录！");
  window.location.href="site.asp?id=<%=session("id")%>";

</SCRIPT>