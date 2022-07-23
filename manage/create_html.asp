<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<!--#include file="inc_class.asp"-->
<meta http-equiv="pragma" content="no-cache"> 
<meta http-equiv="cache-control" content="no-cache"> 
<meta http-equiv="expires" content="0"> 
<%
if SaveFile("../index.html","http://www.3aicn.com/template/index.asp") then 
response.write "首页生成成功"
response.write "<br>"
else 
Response.write "首页没有生成" 
end if 


'===========信息内容页======'
dim rspro,sqlpro,id,nb
sqlpro="select * from news"
Set rspro= Server.CreateObject("ADODB.Recordset")
rspro.open sqlpro,conn,3,2
do while not rspro.eof
product_id=rspro("product_id")
types=rspro("type")
			if SaveFile("../"&product_id&".html","http://www.3aicn.com/template/newsinfo.asp?id="&product_id&"&types="&types&"") then 
			response.write "内容页已生成"
			response.write "<br>"
			else 
			Response.write "内容页没有生成" 
			end if 
rspro.movenext
loop
rspro.close
set rspro=nothing

'===========信息内容页======'


'===========产品内容页======'
sqlpro="select * from product"
Set rspro= Server.CreateObject("ADODB.Recordset")
rspro.open sqlpro,conn,3,2
do while not rspro.eof
product_id=rspro("product_id")
			if SaveFile("../pro_"&product_id&".html","http://www.3aicn.com/template/productshow.asp?productid="&product_id&"") then 
			response.write "产品内容页已生成"
			response.write "<br>"
			else 
			Response.write "产品内容页没有生成" 
			end if 
rspro.movenext
loop
rspro.close
set rspro=nothing

'===========产品内容页======'


'===========企业信息生成======'
sqlpro="select * from company where ver='cn'"
Set rs_com= Server.CreateObject("ADODB.Recordset")
rs_com.open sqlpro,conn,1,1
do while not rs_com.eof
	id=rs_com("id")
	names=rs_com("name")
	if SaveFile("../"&names&".html","http://www.3aicn.com/template/company.asp?id="&id&"") then   
    end if 
	rs_com.movenext
	loop
	rs_com.close
	set rs_com=nothing
	response.Write("企业信息生成") 
'===========页面信息生成======'	
'===========产品大类列表页======'

Set rs_root= Server.CreateObject("ADODB.Recordset")
sql_root="select * from product_cate where root_id=0"
rs_root.open sql_root,conn,1,1
do while not rs_root.eof
root_id=rs_root("PRO_CATE_ID")



	Set rs_big= Server.CreateObject("ADODB.Recordset")
	sql_big="select * from product where BIGTYPE="&root_id&""
	rs_big.open sql_big,conn,1,1

	do while not rs_big.eof
	   PageSize=12
	   PageCount = Int(rs_big.RecordCount / PageSize*-1)*-1
	   if PageCount=0 then PageCount=1 end if
	   
		for k=1 to PageCount
		if SaveFile("../probig_"&root_id&"_"&k&".html","http://www.3aicn.com/template/product_big.asp?bigtype="&root_id&"&page="&k&"") then
		end if 
		next
		k=k+1
			
	
	if SaveFile("../probig_"&root_id&".html","http://www.3aicn.com/template/product_big.asp?bigtype="&root_id&"") then   
	end if 
	rs_big.movenext
	loop
	rs_big.close
	set rs_big=nothing


rs_root.movenext
loop
rs_root.close
set rs_root=nothing
'===========产品大类列表页======'



'===========产品小类列表页======'

'Set rs_parent= Server.CreateObject("ADODB.Recordset")
'sql_parent="select * from product_cate where root_id<>0"
'rs_parent.open sql_parent,conn,1,1
'do while not rs_parent.eof
'parent_id=rs_parent("PRO_CATE_ID")



'	Set rs_small= Server.CreateObject("ADODB.Recordset")
'	sql_small="select * from product where type="&parent_id&""
'	rs_small.open sql_small,conn,1,1

'	do while not rs_small.eof
'	   PageSize=12
'	   PageCount = Int(rs_small.RecordCount / PageSize*-1)*-1
	   
'		for k=1 to PageCount
'		if SaveFile("../prosmall_"&parent_id&"_"&k&".html","http://www.3aicn.com/template/product_small.asp?cateid="&parent_id&"&page="&k&"") then
'		end if 
'		next
'		k=k+1
			
	
'	if SaveFile("../prosmall_"&parent_id&".html","http://www.3aicn.com/template/product_small.asp?cateid="&parent_id&"") then   
'	end if 
'	rs_small.movenext
'	loop
'	rs_small.close
'	set rs_small=nothing


'rs_parent.movenext
'loop
'rs_parent.close
'set rs_parent=nothing
'===========产品小类列表页======'

'===========产品列表页======'

'	Set rs_small= Server.CreateObject("ADODB.Recordset")
'	sql_small="select * from product where ver='cn'"
'	rs_small.open sql_small,conn,1,1

'	do while not rs_small.eof
'	   PageSize=12
'	   PageCount = Int(rs_small.RecordCount / PageSize*-1)*-1
	   
'		for k=1 to PageCount
'		if SaveFile("../product_"&k&".html","http://www.3aicn.com/template/product.asp?page="&k&"") then
'		end if 
'		next
'		k=k+1
			
	
'	if SaveFile("../product.html","http://www.3aicn.com/template/product.asp") then   
'	end if 
'	rs_small.movenext
'	loop
'	rs_small.close
'	set rs_small=nothing

'===========产品列表页======'

'===========信息中心列表页======'
Set rs_news_cate= Server.CreateObject("ADODB.Recordset")
sql_news_cate="select * from news_cate where ver='cn'"
rs_news_cate.open sql_news_cate,conn,1,1
do while not rs_news_cate.eof
news_id=rs_news_cate("id")
newsname=rs_news_cate("filename")    
	Set rs_news= Server.CreateObject("ADODB.Recordset")
	sql_news="select * from news where type="&news_id&""
	rs_news.open sql_news,conn,1,1

	do while not rs_news.eof
	   PageSize=12
	   PageCount = Int(rs_news.RecordCount / PageSize*-1)*-1
	   if PageCount=0 then PageCount=1 end if
		for k=1 to PageCount
		if SaveFile("../"&newsname&"_"&k&".html","http://www.3aicn.com/template/news.asp?type="&news_id&"&page="&k&"") then
		end if 
		next
		k=k+1
			
	
	if SaveFile("../"&newsname&".html","http://www.3aicn.com/template/news.asp?type="&news_id&"") then   
	end if 
	
	rs_news.movenext
	loop
	rs_news.close
	set rs_news=nothing
	
rs_news_cate.movenext
loop
rs_news_cate.close
set rs_news_cate=nothing

'===========信息中心列表页======'

	

function SaveFile(LocalFileName,RemoteFileUrl) 
Dim Ads, Retrieval, GetRemoteData 
On Error Resume Next 
Set Retrieval = Server.CreateObject("Microso" & "ft.XM" & "LHTTP") 
With Retrieval 
.Open "Get", RemoteFileUrl, False, "", "" 
.Send 
GetRemoteData = .ResponseBody 
End With 
Set Retrieval = Nothing 
Set Ads = Server.CreateObject("Ado" & "db.Str" & "eam") 
With Ads 
.Type = 1 
.Open 
.Write GetRemoteData 
.SaveToFile Server.MapPath(LocalFileName), 2 
.Cancel() 
.Close() 
End With 
Set Ads=nothing 
if err <> 0 then 
SaveFile = false 
err.clear 
else 
SaveFile = true 
end if 
End function 
%>
