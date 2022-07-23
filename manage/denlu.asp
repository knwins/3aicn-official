<%
if session("userlogin")="" then
 response.write "<script>alert('ÄãµÇÂ¼ÒÑ³¬Ê±£¡');window.parent.location.href='login.asp';</script>"
end if
%>