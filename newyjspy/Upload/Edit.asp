<!--#include file="Config.asp"-->
<!--#include file="../admin/session.asp"-->
<%
Dim ID,Msg
	ID = Request.QueryString("ID")
	If Request.QueryString("Action") = "Save" Then SaveData ID
	Sub SaveData(ID)
	If ID < 1 Then 
	Response.Write("��������")
	Response.End()
	End If
	myConn.execute("update Fupload set FILETITLE='"&Request.Form("fname")&"',FILEDESC='"&Request.Form("fdesc")&"',FILETYPE='"&Request.Form("ftype")&"',FILEPATH='"&Request.Form("fpath")&"',FILESIZE='"&Request.Form("fsize")&"' where Fupload_ID="&ID)
	Msg = "�ɹ��޸����ļ�������Ϣ"
	End Sub

If msg <> "" Then
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_List.asp'>"&Msg&"<br>��ҳ����3���ڷ���<BR>�����������û�з�Ӧ����<a href=Admin_List.asp>����˴�����</a>")
Response.End()
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ϴ�����</title>
<link href="../style.css" rel="stylesheet" type="text/css">
</head>

<body>
<div align="center">
  <!--#include file="top1.asp" -->
  <table width="840" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="20" rowspan="2" background="../images/leftbk.jpg">&nbsp;</td>
      <td height="40" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;���ĵ�ǰλ�ã�&gt;&gt; <span class="style3"><a href="Admin_main.asp">ϵͳ����</a></span> &gt;&gt; <span class="style3">�༭�ļ�����</span>&nbsp;&nbsp; <a href="../admin/admin_logout.asp">��<span class="style2">ע��</span>��</a></td>
      <td width="20" rowspan="2" background="../images/rightbk.jpg">&nbsp;</td>
    </tr>
    <tr>
      <td><div align="center">
          <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#0153A3">
            <form name="Edit" action="Edit.asp?Action=Save&ID=<%=ID%>" method="post">
              <tr>
                <td height="25"><%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from Fupload Where Fupload_ID="&ID
	frst.open sql,myconn,1,1
	If not frst.Eof then
			fid = frst("Fupload_ID").Value
			ftitle = frst("FileTitle").Value
			fdesc = frst("FileDesc").Value
			ftype = frst("FileType").Value
			fpath = frst("FilePath").Value
			fsize = frst("Filesize").Value
			fuploadtime = frst("uploadTime").Value
'FileNameStr=Server.Mappath(fpath)
'Isize.GetImgSize Cstr(FileNameStr)
	%>
                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
                      <tr class="text">
                        <td width="150" bgcolor="#FFFFFF"><div align="right">�ļ����ƣ�</div></td>
                        <td bgcolor="#FFFFFF"><input type="text" name="fname" class="TextBoxT" value="<%=ftitle%>">
                        </td>
                        <td width="30%" rowspan="3" bgcolor="#FFFFFF">&nbsp;</td>
                      </tr>
                      <tr class="text">
                        <td width="150" bgcolor="#FFFFFF"><div align="right">�ļ����ͣ�</div></td>
                        <td bgcolor="#FFFFFF"><select name="ftype" class="TextBoxT" id="filetype">
                          <option value="��    ��" selected>�� ��</option>
                          <option value="ͼ    Ƭ">ͼ Ƭ</option>
                          <option value="ý    ��">ý ��</option>
                        </select></td>
                      </tr>
                      <tr class="text">
                        <td width="150" bgcolor="#FFFFFF"><div align="right">�ļ�·����</div></td>
                        <td bgcolor="#FFFFFF"><input name="fpath" type="text" class="TextBoxT" value="<%=fpath%>" size="50">
                            <%
			Set Fs = Server.CreateObject("Scripting.FileSystemObject")
			If Fs.FileExists(server.mappath(fPath)) Then
			Response.Write("<img src=Images/isexists.gif")
			End If
			%>
                        </td>
                      </tr>
                      <tr class="text">
                        <td width="150" bgcolor="#FFFFFF"><div align="right">˵����</div></td>
                        <td colspan="2" bgcolor="#FFFFFF"><input type="text" name="fdesc" class="TextBoxT" value="<%=fdesc%>"></td>
                      </tr>
                      <tr class="text">
                        <td height="1" align="right" bgcolor="#FFFFFF">�ļ���С��</td>
                        <td height="1" colspan="2" bgcolor="#FFFFFF"><input type="text" name="fsize" class="TextBoxT" value="<%=fsize%>">
                        bytes</td>
                      </tr>
                      <tr class="text">
                        <td height="1" align="right" bgcolor="#FFFFFF">&nbsp;</td>
                        <td height="1" colspan="2" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="�޸�">
&nbsp;
                        <input type="button" name="Submit2" value="����" onclick='javascript:window.location="Admin_List.asp"'></td>
                      </tr>
                      <tr class="text">
                        <td height="1" align="right" bgcolor="#FFFFFF">&nbsp;</td>
                        <td height="1" colspan="2" bgcolor="#FFFFFF"></td>
                      </tr>
                    </table>
                    <%
	  else
	  %>
                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
                      <tr class="text">
                        <td bgcolor="#FFFFFF">û�ж�Ӧ�����ݣ�</td>
                      </tr>
                    </table>
                    <%
	  end if
	  frst.close
	  set frst = nothing
	  myconn.close
	  set myconn = nothing
	  %>
                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
                      <tr class="text">
                        <td align="right" class="heading" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;</td>
                      </tr>
                  </table></td>
              </tr>
            </form>
          </table>
      </div></td>
    </tr>
  </table>
  <!--#include file="bottom1.asp" -->
</div>
</body>


</html>