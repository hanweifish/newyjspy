<!--#include file="Config.asp"-->
<!--#include file="../admin/session.asp"-->


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ͼƬ�ϴ�</title>
<link href="../style.css" rel="stylesheet" type="text/css">
</head>

<body>
<div align="center">
  <!--#include file="top1.asp" -->
  <table width="840" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="20" rowspan="2" background="../images/leftbk.jpg">&nbsp;</td>
      <td height="40" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;���ĵ�ǰλ�ã�&gt;&gt; <span class="style3">�ϴ��ļ�</span> -- <a href="Admin_List.asp">�ļ�����</a> -- <a href="Admin_main.asp">ϵͳ����</a>&nbsp;&nbsp; <a href="../admin/admin_logout.asp">��<span class="style2">ע��</span>��</a></td>
      <td width="20" rowspan="2" background="../images/rightbk.jpg">&nbsp;</td>
    </tr>
    <tr>
      <td><div align="center">
          <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
            <tr>
              <td height="25">
<%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from Fupload order by Uploadtime desc"
	frst.open sql,myconn,1,1
	fcount = frst.recordcount
	if fcount > 0 then 	
		dim tbbgcolor
		dim maxperpage,pages,page
		maxperpage = 4
		frst.pagesize = maxperpage
		pages = frst.pagecount
		page = Request.QueryString("page")
		if not isnumeric(page) then page = 1 else page = cint(page)
		if page < 1 then page = 1
		if page > pages then page = pages
		frst.absolutepage = page
		for i = 1 to maxperpage
			if frst.eof then exit for
			if i mod 2 = 0 then tbbgcolor = "" else tbbgcolor = "#0066cc"
			fid = frst("Fupload_ID").Value
			ftitle = frst("FileTitle").Value
			fdesc = frst("FileDesc").Value
			ftype = frst("FileType").Value
			fpath = frst("FilePath").Value
			fsize = frst("Filesize").Value
			fuploadtime = frst("uploadTime").Value
	%>
                  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
                    <tr class="text">
                      <td width="150"><div align="right">�ļ����ƣ�</div></td>
                      <td><a href="<%=fpath%>" target="_NEWwIN"><%=ftitle%>( �ļ���С��<%=fsize%> bytes)</a> </td>
                      <td align="right"></td>
                    </tr>
                    <tr class="text">
                      <td width="150"><div align="right">�ļ����ͣ�</div></td>
                      <td colspan="2"><%=ftype%></td>
                    </tr>
                    <tr class="text">
                      <td width="150"><div align="right">�ϴ����ڣ�</div></td>
                      <td colspan="2"><%=fuploadtime%></td>
                    </tr>
                    <tr class="text">
                      <td width="150"><div align="right">˵����</div></td>
                      <td colspan="2"><%=fdesc%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td height="1"></td>
                      <td height="1" colspan="2"></td>
                    </tr>
                  </table>
                  <%
		  	frst.movenext
		next
	  else
	  %>
                  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
                    <tr class="text">
                      <td bgcolor="#FFFFFF">��û���ϴ������ݣ�</td>
                    </tr>
                  </table>
                  <%
	  end if
	  frst.close
	  set frst = nothing
	  myconn.close
	  set myconn = nothing
	  %>
                  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
                    <tr class="text">
                      <td align="center"><%
		  If Page > 1 Then Response.Write ("<a href='?page=1'>��ҳ</a><a href='?page="& Page - 1 &"'>��һҳ</a>")
		  If page < pages Then Response.Write ("&nbsp;<a href='?page="& Page + 1 &"'>��һҳ</a>&nbsp;<a href='?page="& Pages &"'>ĩҳ</a>")
		  %></td>
                    </tr>
                </table></td>
            </tr>
          </table>
          <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#003366">
            <tr>
              <td height="25" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
                  <form action="SaveUpload.asp" method="post" enctype="multipart/form-data" name="form1">
                    <tr bgcolor="#006699" class="text">
                      <td width="33%" height="25" bgcolor="#FFFFFF"><div align="right"><strong>�ϴ����ݣ�</strong></div></td>
                      <td bgcolor="#FFFFFF"><%
		Response.Write "  �����ϴ����ļ�����:<br> "
		Set Fs = Server.CreateObject("Scripting.FileSystemObject")
		For Each str In OKAr
		If Fs.FileExists(Server.MapPath("Images\"& Str &".gif")) Then
		Response.Write "<img src='Images/" & str & ".gif' alt='" & str & "�ļ�'> "
		Else
		Response.Write "<img src='Images/X.gif' alt='" & str & "�ļ�'> "
		End If
		Next
		Set Fs = Nothing
		Response.Write "<br>  �����ϴ����ļ����:"&Oksize / 1024&"KB"
		%></td>
                    </tr>
                    <tr class="text">
                      <td width="33%"><div align="right"><strong>�ļ����ƣ�</strong></div></td>
                      <td><input name="filetitle" type="text" class="TextBoxT" size="25">
                          <select name="filetype" class="TextBoxT" id="filetype">
                            <option value="��    ��" selected>�� ��</option>
                            <option value="ͼ    Ƭ">ͼ Ƭ</option>
                            <option value="ý    ��">ý ��</option>
                          </select>
                      </td>
                    </tr>
                    <tr class="text">
                      <td valign="top"><div align="right"><strong>�ϴ����ļ���</strong></div></td>
                      <td><input name="filedata" type="file" class="TextBoxT" id="filedata" size="27"></td>
                    </tr>
                    <tr class="text">
                      <td valign="top"><div align="right"><strong>�ļ�˵����</strong><br>
                      </div></td>
                      <td><textarea name="filedesc" cols="36" rows="4" class="TextBoxT" id="filedesc"></textarea></td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                      <td><input type="submit" name="Submit" value="�ϴ�����"></td>
                    </tr>
                  </form>
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
