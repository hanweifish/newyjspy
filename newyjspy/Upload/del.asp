<!--#include file="Config.asp"-->
<!--#include file="../admin/session.asp"-->
<%
Dim GetID,ID
GetID = split(Trim(Request.Form("DelID")),",")
Set Fs = Server.CreateObject("Scripting.FileSystemObject")
For each Str in GetID
VarStr = Split(Str,"|")
ID = VarStr(0)
Path = VarStr(1)
myconn.execute("Delete From Fupload Where Fupload_ID="&ID)
If Fs.FileExists(server.mappath(Path)) Then
Set Os = Fs.GetFile(server.mappath(Path))
Os.Delete
Response.Write Path&"已被删除！<br>"
Else
Response.Write Path&"此图片不存在！数据已被删除<br>"
End If
Next
%>

<head>
<meta http-equiv="refresh" content="3;URL=Admin_List.asp">
<link href="../style.css" rel="stylesheet" type="text/css">
</head>

<body >
<div align="center">
<!--#include file = "top1.asp"-->
<p class="smtext">&nbsp;</p>
<p class="smtext">本页将在3秒后返回</p>
<p class="smtext"><br>
    如果您的浏览器没有反应，请<a href=Admin_List.asp><font color=000000><b>点击此处返回</b></font></a>
</p>
</div>
</body>

