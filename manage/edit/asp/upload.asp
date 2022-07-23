<% Option Explicit %>
<%
'######################################
' eWebEditor v3.00 - Advanced online web based WYSIWYG HTML editor.
' Copyright (c) 2003-2004 eWebSoft.com
'
' For further information go to http://www.ewebsoft.com/
' This copyright notice MUST stay intact for use.
'######################################

Session("eWebEditor_Original_CodePage") = Session.CodePage
Session.CodePage = 65001

%>
<!--#include file="config.asp"-->
<!--#include file="upfileclass.asp"-->

<%
Server.ScriptTimeOut = 1800

Dim sType, sStyleName, sLanguage
Dim sAllowExt, nAllowSize, sUploadDir, nUploadObject, nAutoDir, sBaseUrl, sContentPath
Dim sFileExt, sOriginalFileName, sSaveFileName, sPathFileName, nFileNum
Dim nSLTFlag, nSLTMinSize, nSLTOkSize, nSYFlag, sSYText, sSYFontColor, nSYFontSize, sSYFontName, sSYPicPath, nSLTSYObject, sSLTSYExt, nSYMinSize, sSYShadowColor, nSYShadowOffset

Call InitUpload()


Dim sAction
sAction = UCase(Trim(Request.QueryString("action")))

Select Case sAction
Case "REMOTE"
	Call DoCreateNewDir()
	Call DoRemote()
Case "SAVE"
	Call ShowForm()
	Call DoCreateNewDir()
	Call DoSave()
Case Else
	Call ShowForm()
End Select

Session.CodePage = Session("eWebEditor_Original_CodePage")


Sub ShowForm()
%>
<HTML>
<HEAD>
<TITLE>eWebEditor</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript" src="../dialog/dialog.js"></script>
<link href='../language/<%=sLanguage%>.css' type='text/css' rel='stylesheet'>
<link href='../dialog/dialog.css' type='text/css' rel='stylesheet'>
</head>
<body class=upload>

<form action="?action=save&type=<%=sType%>&style=<%=sStyleName%>&language=<%=sLanguage%>" method=post name=myform enctype="multipart/form-data">
<input type=file name=uploadfile size=1 style="width:100%" onchange="originalfile.value=this.value">
<input type=hidden name=originalfile value="">
</form>

<script language=javascript>

var sAllowExt = "<%=sAllowExt%>";

function CheckUploadForm() {
	if (!IsExt(document.myform.uploadfile.value,sAllowExt)){
		parent.UploadError('lang["ErrUploadInvalidExt"]+":'+sAllowExt+'"');
		return false;
	}
	return true
}

var oForm = document.myform ;
oForm.attachEvent("onsubmit", CheckUploadForm) ;
if (! oForm.submitUpload) oForm.submitUpload = new Array() ;
oForm.submitUpload[oForm.submitUpload.length] = CheckUploadForm ;
if (! oForm.originalSubmit) {
	oForm.originalSubmit = oForm.submit ;
	oForm.submit = function() {
		if (this.submitUpload) {
			for (var i = 0 ; i < this.submitUpload.length ; i++) {
				this.submitUpload[i]() ;
			}
		}
		this.originalSubmit() ;
	}
}

try {
	parent.UploadLoaded();
}
catch(e){
}

</script>

</body>
</html>
<% 
End Sub 


Sub DoSave()

	Select Case nUploadObject
	Case 1
		Call DoUpload_ASPUpload()
	Case 2
		Call DoUpload_SAFileUP()
	Case 3
		Call DoUpload_LyfUpload()
	Case Else
		Call DoUpload_Class()
	End Select
	
	Dim s_SmallImageFile, s_SmallImagePathFile, s_SmallImageScript
	s_SmallImageFile = getSmallImageFile(sSaveFileName)
	s_SmallImagePathFile = ""
	s_SmallImageScript = ""
	If makeImageSLT(sUploadDir, sSaveFileName, s_SmallImageFile) = True Then
		Call makeImageSY(sUploadDir, s_SmallImageFile)
		Call makeImageSY(sUploadDir, sSaveFileName)
		s_SmallImagePathFile = sContentPath & s_SmallImageFile
		s_SmallImageScript = "try{obj.addUploadFile('" & sOriginalFileName & "', '" & s_SmallImageFile & "', '" & s_SmallImagePathFile & "');} catch(e){} "
	Else
		s_SmallImageFile = ""
		Call makeImageSY(sUploadDir, sSaveFileName)
	End If

	sPathFileName = sContentPath & sSaveFileName
	Call OutScript("parent.UploadSaved('" & sPathFileName & "','" & s_SmallImagePathFile & "');var obj=parent.dialogArguments.dialogArguments;if (!obj) obj=parent.dialogArguments;try{obj.addUploadFile('" & sOriginalFileName & "', '" & sSaveFileName & "', '" & sPathFileName & "');} catch(e){} " & s_SmallImageScript)

End Sub

Sub makeImageSY(s_Path, s_File)
	If nSYFlag = 0 Then Exit Sub
	If isValidSLTSYExt(s_File) = False Then Exit Sub

	On Error Resume Next
	Dim nOriginalWidth
	Dim oImage, oLogo

	Select Case nSLTSYObject
	Case 0
		If IsObjInstalled("Persits.Jpeg") = False Then Exit Sub
		Set oImage = Server.CreateObject("Persits.Jpeg")
		oImage.Open Server.Mappath(s_Path & s_File)
		nOriginalWidth = oImage.OriginalWidth
		If nSYMinSize > nOriginalWidth Then Exit Sub
		If nSYFlag = 1 Then
			oImage.Canvas.Font.Color = Clng("&H" & sSYFontColor)
			oImage.Canvas.Font.Family = sSYFontName
			'oImage.Canvas.Font.Bold = True
			oImage.Canvas.Font.Size = nSYFontSize
			oImage.Canvas.Font.ShadowColor = Clng("&H" & sSYShadowColor)
			oImage.Canvas.Font.ShadowXOffset = nSYShadowOffset
			oImage.Canvas.Font.ShadowYOffset = nSYShadowOffset
			oImage.Canvas.Print 5, 5, sSYText
			oImage.Save Server.Mappath(s_Path & s_File)
		End If
		If nSYFlag = 2 Then
			Set oLogo = Server.CreateObject("Persits.Jpeg")
			oLogo.Open Server.Mappath(sSYPicPath)
			oImage.DrawImage 0, 0, oLogo
			oImage.Save Server.Mappath(s_Path & s_File)
			Set oLogo = Nothing
		End If
		Set oImage = Nothing
	Case Else

	End Select

End Sub

Function makeImageSLT(s_Path, s_File, s_SmallFile)
	makeImageSLT = False
	If nSLTFlag = 0 Then Exit Function
	If isValidSLTSYExt(s_File) = False Then Exit Function

	Dim nOriginalWidth, nOriginalHeight, nWidth, nHeight
	Dim oImage

	Select Case nSLTSYObject
	Case 0
		If IsObjInstalled("Persits.Jpeg") = False Then Exit Function
		Set oImage = Server.CreateObject("Persits.Jpeg")
		oImage.Open Server.Mappath(s_Path & s_File)
		nOriginalWidth = oImage.OriginalWidth
		nOriginalHeight = oImage.OriginalHeight
		If nOriginalWidth < nSLTMinSize And nOriginalHeight < nSLTMinSize Then Exit Function
		If nOriginalWidth > nOriginalHeight Then
			nWidth = nSLTOkSize
			nHeight = (nSLTOkSize / nOriginalWidth) * nOriginalHeight
		Else
			nHeight = nSLTOkSize
			nWidth = (nSLTOkSize / nOriginalHeight) * nOriginalWidth
		End If
		oImage.Width = nWidth
		oImage.Height = nHeight
		oImage.Save Server.Mappath(s_Path & s_SmallFile)
		Set oImage = Nothing
	Case Else

	End Select

	makeImageSLT = True
End Function

Function isValidSLTSYExt(s_File)
	Dim b, i, aExt, sExt
	b = False
	sExt = LCase(Mid(s_File, InstrRev(s_File, ".")+1))
	aExt = Split(LCase(sSLTSYExt), "|")
	For i = 0 To UBound(aExt)
		If aExt(i) = sExt Then
			b = True
			Exit For
		End If
	Next
	isValidSLTSYExt = b
End Function

Function getSmallImageFile(s_File)
	Dim n
	n = InstrRev(s_File, ".")
	getSmallImageFile = Left(s_File, n-1) & "_s." & Mid(s_File, n+1)
End Function


Sub DoRemote()
	Dim sContent, i
	For i = 1 To Request.Form("eWebEditor_UploadText").Count 
		sContent = sContent & Request.Form("eWebEditor_UploadText")(i) 
	Next
	If sAllowExt <> "" Then
		sContent = ReplaceRemoteUrl(sContent, sAllowExt)
	End If

	Response.Write "<HTML><HEAD><TITLE>eWebEditor</TITLE><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head><body>" & _
		"<input type=hidden id=UploadText value=""" & inHTML(sContent) & """>" & _
		"</body></html>"

	Call OutScriptNoBack("parent.setHTML(UploadText.value);try{parent.addUploadFile('" & sOriginalFileName & "', '" & sSaveFileName & "', '" & sPathFileName & "');} catch(e){} parent.remoteUploadOK();")

End Sub


Sub DoCreateNewDir()
	If IsObjInstalled("Scripting.FileSystemObject") = False Then
		Exit Sub
	End If

	Dim sCreateDir
	Select Case nAutoDir
	Case 1
		sCreateDir = Left(FormatTime(Now(), 4), 4)
	Case 2
		sCreateDir = Left(FormatTime(Now(), 4), 6)
	Case 3
		sCreateDir = Left(FormatTime(Now(), 4), 8)
	Case Else
		Exit Sub
	End Select
	sUploadDir = sUploadDir & sCreateDir & "/"
	sContentPath = sContentPath & sCreateDir & "/"

	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(Server.Mappath(sUploadDir)) = False Then
		fso.CreateFolder Server.Mappath(sUploadDir)
	End If
	Set fso = Nothing
End Sub

Sub DoUpload_LyfUpload()
	On Error Resume Next
	Dim oUpload, sResult, sOriginalFile
	Set oUpload = Server.CreateObject("LyfUpload.UploadFile")
	oUpload.CodePage = 65001
	oUpload.ExtName = Replace(sAllowExt, "|", ",")
	oUpload.MaxSize = nAllowSize*1024
	sOriginalFile = oUpload.Request("originalfile")
	sOriginalFileName = Mid(sOriginalFile, InStrRev(sOriginalFile, "\") + 1)
	sFileExt = LCase(Mid(sOriginalFileName, InStrRev(sOriginalFileName, ".") + 1))
	Call CheckValidExt(sFileExt)
	sSaveFileName = GetRndFileName(sFileExt)
	sResult = oUpload.SaveFile("uploadfile", Server.Mappath(sUploadDir), True, sSaveFileName)

	Select Case sResult
	Case "0"
		Call OutScript("parent.UploadError('lang[""ErrUploadSizeLimit""]+"":" & nAllowSize & "KB""')")
	Case ""
		Call OutScript("parent.UploadError('lang[""ErrUploadInvalidFile""]')")
	Case "1"
		Call OutScript("parent.UploadError('lang[""ErrUploadInvalidExt""]+"":" & sAllowExt & """')")
	End Select
	
	Set oUpload = Nothing
End Sub

Sub DoUpload_SAFileUp()
	On Error Resume Next
	Dim oFileUp
	Set oFileUp = Server.CreateObject("SoftArtisans.FileUp")
	'oFileUp.MaxBytes = nAllowSize*1024
	oFileUp.CodePage = 65001
	oFileUp.Path = Server.MapPath(sUploadDir)

	If oFileUp.Form("uploadfile").TotalBytes > nAllowSize*1024 Then
		Err.Clear
		Call OutScript("parent.UploadError('lang[""ErrUploadSizeLimit""]+"":" & nAllowSize & "KB""')")
	End If
	If oFileUp.Form("uploadfile").IsEmpty Then
		Call OutScript("parent.UploadError('lang[""ErrUploadInvalidFile""]')")
	End If

	Dim sShortFileName
	'sShortFileName = oFileUp.Form("uploadfile").ShortFileName
	sShortFileName = Mid(oFileUp.Form("uploadfile").UserFilename, InstrRev(oFileUp.Form("uploadfile").UserFilename, "\") + 1)

	sFileExt = LCase(Mid(sShortFileName, InStrRev(sShortFileName, ".") + 1))
	Call CheckValidExt(sFileExt)
	sOriginalFileName = sShortFileName
	sSaveFileName = GetRndFileName(sFileExt)
	oFileUp.Form("uploadfile").SaveAs Server.Mappath(sUploadDir & sSaveFileName)
	
	Set oFileUp = Nothing
End Sub

Sub DoUpload_ASPUpload()
	On Error Resume Next
	Dim oUpload, oFile, nCount
	Set oUpload = Server.CreateObject("Persits.Upload")
	oUpload.CodePage = 65001
	oUpload.SetMaxSize nAllowSize*1024, True
	nCount = oUpload.Save
	'nCount = oUpload.SaveToMemory

	If nCount < 1 Then
		Call OutScript("parent.UploadError('lang[""ErrUploadInvalidFile""]')")
	End If
	If Err.Number = 8 Then
		Err.Clear
		Call OutScript("parent.UploadError('lang[""ErrUploadSizeLimit""]+"":" & nAllowSize & "KB""')")
	End If
	
	Set oFile = oUpload.Files("uploadfile")
	sFileExt = LCase(Mid(oFile.Ext, 2))
	Call CheckValidExt(sFileExt)
	sOriginalFileName = oFile.FileName
	sSaveFileName = GetRndFileName(sFileExt)
	oFile.SaveAs Server.Mappath(sUploadDir & sSaveFileName)

	Set oFile = Nothing
	Set oUpload = Nothing
End Sub

Sub DoUpload_Class()
	On Error Resume Next
	Dim oUpload, oFile
	Set oUpload = New upfile_class

	oUpload.GetData nAllowSize*1024

	If oUpload.Err > 0 Then
		Select Case oUpload.Err
		Case 1
			Call OutScript("parent.UploadError('lang[""ErrUploadInvalidFile""]')")
		Case 2
			Call OutScript("parent.UploadError('lang[""ErrUploadSizeLimit""]+"":" & nAllowSize & "KB""')")
		End Select
	End If

	Set oFile = oUpload.File("uploadfile")
	sFileExt = LCase(oFile.FileExt)
	Call CheckValidExt(sFileExt)
	sOriginalFileName = oFile.FileName
	sSaveFileName = GetRndFileName(sFileExt)

	Dim str_Mappath
	str_Mappath = Server.Mappath(sUploadDir & sSaveFileName)
	sFileExt = LCase(Mid(str_Mappath, InstrRev(str_Mappath, ".") + 1))
	Call CheckValidExt(sFileExt)

	oFile.SaveToFile str_Mappath
	
	Set oFile = Nothing
	Set oUpload = Nothing
End Sub

Function GetRndFileName(sExt)
	Dim sRnd
	Randomize
	sRnd = Int(900 * Rnd) + 100
	GetRndFileName = year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & sRnd & "." & sExt
End Function

Sub OutScript(str)
	Response.Write "<script language=javascript>" & str & ";history.back()</script>"
	Session.CodePage = Session("eWebEditor_Original_CodePage")
	Response.End
End Sub
Sub OutScriptNoBack(str)
	Response.Write "<script language=javascript>" & str & "</script>"
End Sub

Sub CheckValidExt(sExt)
	Dim b, i, aExt
	b = False
	aExt = Split(sAllowExt, "|")
	For i = 0 To UBound(aExt)
		If LCase(aExt(i)) = sExt Then
			b = True
			Exit For
		End If
	Next
	If b = False Then
		Call OutScript("parent.UploadError('lang[""ErrUploadInvalidExt""]+"":" & sAllowExt & """')")
	End If
End Sub

Sub InitUpload()
	sType = UCase(Trim(Request.QueryString("type")))
	sStyleName = Trim(Request.QueryString("style"))
	sLanguage = Trim(Request.QueryString("language"))

	Dim i, aStyleConfig, bValidStyle
	bValidStyle = False
	For i = 1 To Ubound(aStyle)
		aStyleConfig = Split(aStyle(i), "|||")
		If Lcase(sStyleName) = Lcase(aStyleConfig(0)) Then
			bValidStyle = True
			Exit For
		End If
	Next

	If bValidStyle = False Then
		OutScript("parent.UploadError('lang[""ErrInvalidStyle""]')")
	End If

	sBaseUrl = aStyleConfig(19)
	nUploadObject = Clng(aStyleConfig(20))
	nAutoDir = CLng(aStyleConfig(21))

	sUploadDir = aStyleConfig(3)
	If Left(sUploadDir, 1) <> "/" Then
		sUploadDir = "../" & sUploadDir
	End If

	Select Case sBaseUrl
	Case "0"
		sContentPath = aStyleConfig(23)
	Case "1"
		sContentPath = RelativePath2RootPath(sUploadDir)
	Case "2"
		sContentPath = RootPath2DomainPath(RelativePath2RootPath(sUploadDir))
	End Select

	Select Case sType
	Case "REMOTE"
		sAllowExt = aStyleConfig(10)
		nAllowSize = Clng(aStyleConfig(15))
	Case "FILE"
		sAllowExt = aStyleConfig(6)
		nAllowSize = Clng(aStyleConfig(11))
	Case "MEDIA"
		sAllowExt = aStyleConfig(9)
		nAllowSize = Clng(aStyleConfig(14))
	Case "FLASH"
		sAllowExt = aStyleConfig(7)
		nAllowSize = Clng(aStyleConfig(12))
	Case Else
		sAllowExt = aStyleConfig(8)
		nAllowSize = Clng(aStyleConfig(13))
	End Select

	nSLTFlag = Clng(aStyleConfig(29))
	nSLTMinSize = Clng(aStyleConfig(30))
	nSLTOkSize = Clng(aStyleConfig(31))
	nSYFlag = Clng(aStyleConfig(32))
	sSYText = aStyleConfig(33)
	sSYFontColor = aStyleConfig(34)
	nSYFontSize = Clng(aStyleConfig(35))
	sSYFontName = aStyleConfig(36)
	sSYPicPath = aStyleConfig(37)
	nSLTSYObject = Clng(aStyleConfig(38))
	sSLTSYExt = aStyleConfig(39)
	nSYMinSize = Clng(aStyleConfig(40))
	sSYShadowColor = aStyleConfig(41)
	nSYShadowOffset = Clng(aStyleConfig(42))

End Sub

Function RelativePath2RootPath(url)
	Dim sTempUrl
	sTempUrl = url
	If Left(sTempUrl, 1) = "/" Then
		RelativePath2RootPath = sTempUrl
		Exit Function
	End If

	Dim sWebEditorPath
	sWebEditorPath = Request.ServerVariables("SCRIPT_NAME")
	sWebEditorPath = Left(sWebEditorPath, InstrRev(sWebEditorPath, "/") - 1)
	Do While Left(sTempUrl, 3) = "../"
		sTempUrl = Mid(sTempUrl, 4)
		sWebEditorPath = Left(sWebEditorPath, InstrRev(sWebEditorPath, "/") - 1)
	Loop
	RelativePath2RootPath = sWebEditorPath & "/" & sTempUrl
End Function

Function RootPath2DomainPath(url)
	Dim sHost, sPort
	sHost = Split(Request.ServerVariables("SERVER_PROTOCOL"), "/")(0) & "://" & Request.ServerVariables("HTTP_HOST")
	sPort = Request.ServerVariables("SERVER_PORT")
	If sPort <> "80" Then
		sHost = sHost & ":" & sPort
	End If
	RootPath2DomainPath = sHost & url
End Function

Function ReplaceRemoteUrl(sHTML, sExt)
	Dim s_Content
	s_Content = sHTML
	If IsObjInstalled("Microsoft.XMLHTTP") = False then
		ReplaceRemoteUrl = s_Content
		Exit Function
	End If
	
	Dim re, RemoteFile, RemoteFileurl, SaveFileName, SaveFileType
	Set re = new RegExp
	re.IgnoreCase  = True
	re.Global = True
	re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}(([A-Za-z0-9_-])+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})([^ \f\n\r\t\v\""\'\>]*\/)(([^ \f\n\r\t\v\""\'\>])+[.]{1}(" & sExt & ")))"

	Set RemoteFile = re.Execute(s_Content)
	Dim a_RemoteUrl(), n, i, bRepeat
	n = 0
	' to no repeat array
	For Each RemoteFileurl in RemoteFile
		If n = 0 Then
			n = n + 1
			Redim a_RemoteUrl(n)
			a_RemoteUrl(n) = RemoteFileurl
		Else
			bRepeat = False
			For i = 1 To UBound(a_RemoteUrl)
				If UCase(RemoteFileurl) = UCase(a_RemoteUrl(i)) Then
					bRepeat = True
					Exit For
				End If
			Next
			If bRepeat = False Then
				n = n + 1
				Redim Preserve a_RemoteUrl(n)
				a_RemoteUrl(n) = RemoteFileurl
			End If
		End If		
	Next
	' start replace
	nFileNum = 0
	For i = 1 To n
		SaveFileType = Mid(a_RemoteUrl(i), InstrRev(a_RemoteUrl(i), ".") + 1)
		SaveFileName = GetRndFileName(SaveFileType)
		If SaveRemoteFile(SaveFileName, a_RemoteUrl(i)) = True Then
			nFileNum = nFileNum + 1
			If nFileNum > 0 Then
				sOriginalFileName = sOriginalFileName & "|"
				sSaveFileName = sSaveFileName & "|"
				sPathFileName = sPathFileName & "|"
			End If
			sOriginalFileName = sOriginalFileName & Mid(a_RemoteUrl(i), InstrRev(a_RemoteUrl(i), "/") + 1)
			sSaveFileName = sSaveFileName & SaveFileName
			sPathFileName = sPathFileName & sContentPath & SaveFileName
			s_Content = Replace(s_Content, a_RemoteUrl(i), sContentPath & SaveFileName, 1, -1, 1)
		End If
	Next

	ReplaceRemoteUrl = s_Content
End Function

Function SaveRemoteFile(s_LocalFileName, s_RemoteFileUrl)
	Dim Ads, Retrieval, GetRemoteData
	Dim bError
	bError = False
	SaveRemoteFile = False
	On Error Resume Next
	Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
	With Retrieval
		.Open "Get", s_RemoteFileUrl, False, "", ""
		.Send
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing

	If LenB(GetRemoteData) > nAllowSize*1024 Then
		bError = True
	Else
		Set Ads = Server.CreateObject("Adodb.Stream")
		With Ads
			.Type = 1
			.Open
			.Write GetRemoteData
			.SaveToFile Server.MapPath(sUploadDir & s_LocalFileName), 2
			.Cancel()
			.Close()
		End With
		Set Ads=nothing
	End If

	If Err.Number = 0 And bError = False Then
		SaveRemoteFile = True
	Else
		Err.Clear
	End If
End Function

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

Function inHTML(str)
	Dim sTemp
	sTemp = str
	inHTML = ""
	If IsNull(sTemp) = True Then
		Exit Function
	End If
	sTemp = Replace(sTemp, "&", "&amp;")
	sTemp = Replace(sTemp, "<", "&lt;")
	sTemp = Replace(sTemp, ">", "&gt;")
	sTemp = Replace(sTemp, Chr(34), "&quot;")
	inHTML = sTemp
End Function

Function FormatTime(s_Time, n_Flag)
	Dim y, m, d, h, mi, s
	FormatTime = ""
	If IsDate(s_Time) = False Then Exit Function
	y = cstr(year(s_Time))
	m = cstr(month(s_Time))
	If len(m) = 1 Then m = "0" & m
	d = cstr(day(s_Time))
	If len(d) = 1 Then d = "0" & d
	h = cstr(hour(s_Time))
	If len(h) = 1 Then h = "0" & h
	mi = cstr(minute(s_Time))
	If len(mi) = 1 Then mi = "0" & mi
	s = cstr(second(s_Time))
	If len(s) = 1 Then s = "0" & s
	Select Case n_Flag
	Case 1
		' yyyy-mm-dd hh:mm:ss
		FormatTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
	Case 2
		' yyyy-mm-dd
		FormatTime = y & "-" & m & "-" & d
	Case 3
		' hh:mm:ss
		FormatTime = h & ":" & mi & ":" & s
	Case 4
		' yyyymmdd
		FormatTime = y & m & d
	End Select
End Function
%>