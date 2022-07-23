<!--#include file="upload_wj.inc"-->

<%
set upload=new upload_file
if upload.form("act")="uploadfile" then
	filepath=trim(upload.form("filepath"))
	filelx=trim(upload.form("filelx"))
	
	i=0
	for each formName in upload.File
		set file=upload.File(formName)
 
 fileExt=lcase(file.FileExt)	'得到的文件扩展名不含有.
 if file.filesize<100 then
 	response.write "<span style=""font-family: 宋体; font-size: 9pt"">请先选择你要上传的文件！　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
	response.end
 end if
 if (filelx<>"jpg") then 
 	response.write "<span style=""font-family: 宋体; font-size: 9pt"">该文件类型不能上传！　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
	response.end
 end if
 if filelx="jpg" then
	if fileext<>"gif" and fileext<>"jpg" and fileext<>"bmp"  then
		response.write "<span style=""font-family: 宋体; font-size: 9pt"">只能上传jpg或gif或bmp格式的图片！　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
		response.end
     	end if
 end if
 if filelx="jpg" then
	if file.filesize>(8000*1024) then
		response.write "<span style=""font-family: 宋体; font-size: 9pt"">最大只能上传 800K 的图片文件！　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
		response.end
	end if
 end if

 randomize
 ranNum=int(100*rnd)+100
 filename=filepath&"pro_"&year(now)&month(now)&second(now)&ranNum&"."&fileExt
%>
<%
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  'file.SaveAs Server.mappath(filename)   ''保存文件
  file.SaveToFile Server.mappath(FileName)
  
  'response.Write("<script language='javascript'>")
  'response.write("window.alert('文件上传成功!请不要修改生成的链接地址！');")
  'response.Write("window.close();")
  'response.Write("/script>")
  response.write "          "&file.FileName&"　　上传成功!　　<br>"
  response.write "          新文件名："&FileName&"<br>"
  response.write "          新文件名已复制到所需的位置，可<a href='javascript:window.close();'><font color=red>关闭窗口</font></a>！"
  response.write "<script>window.opener.document."&upload.form("FormName")&"."&upload.form("EditName")&".value='"&FileName&"'</script>"
%>
<%
 end if
 set file=nothing
	next
	set upload=nothing
end if
%>

