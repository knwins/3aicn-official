<%
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