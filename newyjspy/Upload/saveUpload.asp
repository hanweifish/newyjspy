<!--#include file="Config.asp"-->
<!--#include file="../admin/session.asp"-->
<%
Function ChInt( strVal )
Dim vNum
On Error Resume Next
ChInt = 0
If IsNull(strVal) Then
ChInt = 0
Else
If IsNumeric(strVal) Then
ChInt = CLng(strVal)
If Err.Number<>0 Then
Response.Write "<HR>OVFLW:" & strVal & "<HR>"
End If
Else
ChInt = 0
End If
End If
On Error Goto 0
End Function
%>
<%
''������ʼ����

dim formsize,formdata,Msg
formsize = Request.TotalBytes
formdata = Request.BinaryRead(formsize)
UploadSize=True

If formsize = 0 or Formsize > OKsize Then
UploadSize=False
Response.Write"��Ҫ�ϴ����ļ���С������������,��<a href=index.asp>����</a>�޸�����"
Response.End
End If


dim sinfo_Stream
Set Sinfo_Stream = Server.CreateObject("adodb.stream")
Sinfo_Stream.Type = 1		''2������
Sinfo_Stream.Mode = 3		''��дģʽ
Sinfo_Stream.Open
Sinfo_Stream.Write formdata		''������������ݵ�������
''�������ݱ���
dim VbEnter
dim spStr,lenOfspStr,bpos
dim loopcnt,exitflag,ppoint,npoint
''�������ݱ���		
dim FldData,fldHeadStr,infldpos
dim databpos,datalen
dim FldInfo(15,1)
''fldInfo(0)����ͷ����
''fldInfo(1)��������

VbEnter = chrb(13)&chrb(10)''��ȡ��һ��VbEnterλ��
bpos = Instrb(formdata,VbEnter)
SpStr = midb(formdata,1,bpos+1) ''������һ��0d0a
LenOfspStr = lenb(Spstr) 
ppoint = LenOfspStr+1 ''λ��ָ��,ָ��ÿһ�����������ݵĿ�ʼλ��
formdata = midb(formdata,ppoint)
loopcnt = 0   ''����Ԫ��
do 
	bpos = instrb(formdata,spStr) ''�ָ�λ��
	npoint = (ppoint+bpos+lenofspstr-1)  ''ָ����һ������ʼλ��
	if bpos < 1 then
		fldData = midb(formdata,1,instrb(formdata,leftb(spStr,lenOfspstr-2))-1)
		bpos = lenb(fldData)+1
		exitflag = true
	else
		FldData = leftb(formdata,bpos-1)		
		formdata = midb(formdata,bpos+LenOfspstr)
	end if
	infldpos = instrb(fldData,vbEnter&vbEnter)
	fldHeadStr = bytes2bstr(midb(fldData,1,infldpos-1))
	fldInfo(loopcnt,0) = fldHeadStr	''����ͷ
	''Response.Write fldHeadStr&"<br>"
	databpos = (ppoint+infldpos-1+4)
	Sinfo_Stream.Position = databpos-1
	datalen = (bpos-infldpos-6)
	if datalen = 0 then
		fldInfo(loopcnt,1) = ""
	else
		fldInfo(loopcnt,1) = Sinfo_Stream.Read(datalen)
	end if
	ppoint = npoint
	loopcnt = loopcnt + 1
loop until exitflag = true
Sinfo_Stream.close
Set Sinfo_Stream = Nothing


''���ϳ������ݴ�������
''д�����ݿⲢ�����ļ��ϴ���ʼ
Sub SaveData()

ftitle = MyRequest("filetitle")
Msg = ""
		if ftitle = "" then 
			Msg = Msg & "�ļ����ƣ���<br>"
		else
			Msg = Msg & "�ļ����ƣ�"&ftitle&"<br>"
		end if
		ftype = myrequest("fileType")		
		Msg = Msg & "�ļ����ͣ�"&ftype&"<br>"
		filedata = myrequest("filedata")
		filesize = lenb(filedata)
		if  filesize = 0 then 
			Msg = Msg & "�ϴ��ļ���û��<br>"
		else 
			filename = GetFileName("filedata")
			''���Ƽ�������� *.asp
			file_ctype = GetContentType("filedata")
			Msg = Msg & "�ϴ��ļ���"&filename&"&nbsp;&nbsp;&nbsp;"
			Msg = Msg & "��������"&file_ctype&"&nbsp;&nbsp;&nbsp;"
			Msg = Msg & "�ļ����ȣ�"&filesize&"<br>"
		end if
		filedesc = myrequest("fileDesc")
		Msg = Msg & "�ļ�˵����"&filedesc&"<br><br>"
		
		FileTypeName = GetFileTypeName(FileName)
		If  IsvalidFile(FileTypeName)=False Then
		Msg = "�ļ����ͷǷ����������ϴ�"&FileTypeName&"�ļ���"
		Exit Sub
		End If		

		if ftitle<>"" and fileSize > 0 and UploadSize=True then
			''�������ݵ����ݿ�
			dim basepath,sql
			basepath = "./uploadfile/"
			sql = "insert into Fupload (FileTitle,FileDesc,FileType,FilePath,FileSize) values ('"
			sql = sql & ftitle &"','"&filedesc&"','"&ftype&"','"&basepath&filename&"','"&filesize&"')"
			myconn.Execute(sql)
			Call SavetoFile(filedata,basepath,filename)
			Msg = Msg & "�ļ��Ѿ��ϴ�<br>"
		else
			Msg = Msg & "�ϴ�ʧ�ܣ� "&ErrorMsg&"<br>"
		end if
		myconn.close()
		set myconn = nothing

End Sub	
''�ļ��ϴ��Ѿ�д��������ϣ���ʾ��Ϣ����Ϊ����msg
SaveData
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ϴ��ɹ�</title>
<meta http-equiv="refresh" content="10;URL=Index.asp">
<link href="../style.css" rel="stylesheet" type="text/css">
</head>

<body>
<div align="center">
  <!--#include file="top1.asp" -->
  <table width="840" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="20" rowspan="2" background="../images/leftbk.jpg">&nbsp;</td>
      <td height="40" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;���ĵ�ǰλ�ã�&gt;&gt; <a href="Index.asp">�ϴ��ļ�</a> &gt;&gt; ����ͼƬ
&nbsp;&nbsp; <a href="../admin/admin_logout.asp">��ע����</a></td>
      <td width="20" rowspan="2" background="../images/rightbk.jpg">&nbsp;</td>
    </tr>
    <tr>
      <td><div align="center">
          <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="150"><table width="98%" border="0" cellspacing="0" cellpadding="3" bgcolor="#FFFFFF">
                  <tr class="text">
                    <td><strong>��ʾ��Ϣ��</strong></td>
                  </tr>
                  <tr class="text">
                    <td><%=msg%> </td>
                  </tr>
                  <tr class="text">
                    <td><div align="center">��ҳ����10��󷵻� ������������û�з�Ӧ������<a href="Index.asp">�˴�����</a></div></td>
                  </tr>
              </table></td>
            </tr>
          </table>
      </div></td>
    </tr>
  </table>
  <!--#include file="bottom1.asp" -->
</div>
</body>


</html>