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
Response.Write Path&"�ѱ�ɾ����<br>"
Else
Response.Write Path&"��ͼƬ�����ڣ������ѱ�ɾ��<br>"
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
<p class="smtext">��ҳ����3��󷵻�</p>
<p class="smtext"><br>
    ������������û�з�Ӧ����<a href=Admin_List.asp><font color=000000><b>����˴�����</b></font></a>
</p>
</div>
</body>

