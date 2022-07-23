<%
  ' on Error Resume Next
  '*****************禁止单独访问本页面*********************
    If InStr(Lcase(Request.ServerVariables("SCRIPT_NAME")),"conn.asp") Then
	   Response.Write "禁止单独访问本页面"
	   Response.End
    End If
  '*****************禁止单独访问本页面*********************

  '打开数据库
   Dim Conn
   Dim ConnStr
   Dim DBPath
       DBPath  = Server.MapPath("qianwei/qianwei2009.mdb")
   Set Conn    = Server.CreateObject("ADODB.Connection")
       ConnStr = "Provider=Microsoft.Jet.Oledb.4.0; Data Source = "& DBPath &""
	  
	   Conn.Open ConnStr
	 
	   If Err Then
          Err.Clear
          Set Conn = Nothing
          Response.Write "数据库连接出错，请检查连接字串。"
          Response.End
       End If
%>