<!--#include file="conn.asp"-->
<!--#include file="Config.asp"-->
<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("admin_account")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
dim admin_account
admin_account=session("admin_account")
%>

<%
Dim ID,Msg
	ID = Request.QueryString("ID")
	If Request.QueryString("Action") = "Save" Then SaveData ID
	Sub SaveData(ID)
	If ID < 1 Then 
	Response.Write("��������")
	Response.End()
	End If
	myConn.execute("update info1 set FILETITLE='"&Request.Form("fname")&"',FILEDESC='"&Request.Form("fdesc")&"',FILETYPE='"&Request.Form("ftype")&"',FILEPATH='"&Request.Form("fpath")&"',FILESIZE='"&Request.Form("fsize")&"' where ID="&ID)
	Msg = "�ɹ��޸����ļ�������Ϣ"
	End Sub

If msg <> "" Then
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_List.asp'>"&Msg&"<br>��ҳ����3���ڷ���<BR>�����������û�з�Ӧ����<a href=Admin_List.asp>����˴�����</a>")
Response.End()
End If
%>

<html>
<head>
<script language="javascript">
window.status="��ӭ�����Ͼ���ѧ����ϵ�о���������Ϣϵͳ��"
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�о�����Ϣ����ϵͳ</title>
<link href="../../style.css" rel="stylesheet" type="text/css">
<!--#include file="top1.asp"-->
<table width="800"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="187" height="12"></td>
    <td></td>
  </tr>
  <tr>
    <td rowspan="3" valign="top"><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="47" background="../adminimages/adminlogin.gif">&nbsp;</td>
      </tr>
      <tr>
        <td><div align="center">
          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="100"><div align="center">
                <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" background="../../indeximages/loginbk.gif">
                  <tr>
                    <td width="20%"><div align="center">                      </div>
                         <div align="center"></div></td>
                    <td width="60%" class="style3"><%=admin_account%>��</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td ><div align="center">                      </div>
                         <div align="center"></div></td>
                    <td height="30" class="style2">���Ѿ�<span class="style2">��¼�ɹ�</span>,��</td>
                         <td>&nbsp;</td>
                   </tr>
                  <tr>
                    <td><div align="center"></div></td>
                    <td height="30" class="style2">ѡ������Ҫ�ķ���!</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td height="10" colspan="2"><div align="center"></div></td>
                    <td width="20%">&nbsp;</td>
                  </tr>
                </table>
                </div></td>
            </tr>
            <tr>
              <td height="6" background="../../indeximages/loginbk.gif"><div align="center"><img src="../../indeximages/loginbar.gif" width="129" height="2"></div></td>
            </tr>
            <tr>
              <td height="60" valign="center" background="../../indeximages/loginbk.gif"><div align="center"><a href="../admin_logout.asp"><img src="../../includeimages/logout.gif" width="60" height="24" border="0"></a></div></td>
            </tr>
          </table>
        </div></td>
      </tr>
      <tr>
        <td height="77" background="../../indeximages/links.gif">&nbsp;</td>
      </tr>
      <tr>
        <td background="../../indeximages/loginbk.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="35"><div align="center"><a href="http://www.nju.edu.cn/">�� �� �� ѧ</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="http://physics.nju.edu.cn/">�� �� �� �� ϵ</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="http://bbs.nju.edu.cn/">�� �� С �� ��</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="http://grawww.nju.edu.cn/">�� �� �� Ժ</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="http://job.nju.edu.cn/">�� ҵ ָ �� �� ��</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="../../links.asp" class="style3">&gt;&gt;&gt; More</a></div></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="34" background="../../indeximages/loginbottom.gif">&nbsp;</td>
      </tr>
    </table></td>
    <td valign="top"><div align="right">
      <table width="603"  border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="406" height="10">&nbsp;</td>
            <td width="406" background="../../indeximages/midLinkTop.gif">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td colspan="3"><div align="center">
              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="54" background="../../user/userimages/titlebk1.gif"><div align="center"><img src="../../user/userimages/picUpload.gif" width="523" height="45"></div></td>
                  </tr>
                  <tr>
                    <td height="53" background="../../user/userimages/titlebk2.gif">&nbsp;</td>
                  </tr>
                  <tr>
                    <td background="../../user/userimages/titlebk.gif"><div align="center">
                      <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                        <form name="Edit" action="Edit.asp?Action=Save&ID=<%=ID%>" method="post">
                          <tr>
                            <td height="25"><%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from info1 Where Id="&ID
	frst.open sql,myconn,1,1
	If not frst.Eof then
			fid = frst("id").Value
			ftitle = frst("fileTitle").Value
			fdesc = frst("fileDesc").Value
			ftype = frst("fileType").Value
			fpath = frst("filePath").Value
			fsize = frst("filesize").Value
			fhits = frst("hits").Value
			fuploadtime = frst("uploadTime").Value
'FileNameStr=Server.Mappath(fpath)
'Isize.GetImgSize Cstr(FileNameStr)
	%>
                                <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
                                  <tr class="text">
                                    <td width="150"><div align="right">�ļ����ƣ�</div></td>
                                    <td><input type="text" name="fname" class="style3" value="<%=ftitle%>">
                                    </td>
                                    <td rowspan="3">&nbsp;</td>
                                  </tr>
                                  <tr class="text">
                                    <td width="150"><div align="right">�ļ����ͣ�</div></td>
                                    <td><select name="ftype" class="style3" id="filetype">
                                        <option value="�ز�ͼƬ"<%if ftype="�ز�ͼƬ" then%> selected<% end if %>>�ز�ͼƬ</option>
                                        <option value="���ù���"<%if ftype="���ù���" then%> selected<% end if %>>���ù���</option>
                                        <option value="����Դ��"<%if ftype="����Դ��" then%> selected<% end if %>>����Դ��</option>
                                    </select></td>
                                  </tr>
                                  <tr class="text">
                                    <td width="150"><div align="right">�ļ�·����</div></td>
                                    <td><input name="fpath" type="text" class="style3" value="<%=fpath%>" size="50">
                                        <%
			Set Fs = Server.CreateObject("Scripting.FileSystemObject")
			If Fs.FileExists(server.mappath(fPath)) Then
			Response.Write("<img src=Images/isexists.gif")
			End If
			%>
                                    </td>
                                  </tr>
                                  <tr class="text">
                                    <td width="150"><div align="right">˵����</div></td>
                                    <td colspan="2"><input type="text" name="fdesc" class="style3" value="<%=fdesc%>"></td>
                                  </tr>
                                  <tr class="text">
                                    <td height="1" align="right">�ļ���С��</td>
                                    <td height="1" colspan="2"><input type="text" name="fsize" class="style3" value="<%=fsize%>">
              bytes</td>
                                  </tr>
                                  <tr class="text">
                                    <td height="1" colspan="3" align="right"><div align="center">
  <input name="Submit" type="submit" class="style2" value="�޸�">
&nbsp;
  <input name="Submit2" type="button" class="style2" onClick='javascript:window.location="Admin_List.asp"' value="����">
                                    </div></td>
                                    </tr>
                                  <tr class="text">
                                    <td height="1" align="right">&nbsp;</td>
                                    <td height="1" colspan="2"></td>
                                  </tr>
                                </table>
                                <%
	  else
	  %>
                                <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
                                  <tr class="text">
                                    <td>û�ж�Ӧ�����ݣ�</td>
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
                                    <td align="right" class="heading">&nbsp;&nbsp;&nbsp;&nbsp;</td>
                                  </tr>
                              </table></td>
                          </tr>
                        </form>
                      </table>
</div></td>
                  </tr>
                  <tr>
                    <td height="34" background="../../user/userimages/titlebk3.gif">&nbsp;</td>
                  </tr>
                  </table>
            </div></td>
            </tr>
          </table>
    </div></td>
  </tr>
  <tr>
    <td height="15" valign="top">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top">
      <div align="right">
        <!--#include file="server.asp"-->        

      </div></td>
  </tr>
  <tr>
    <td height="12"></td>
    <td></td>
  </tr>
</table>
<!--#include file="bottom1.asp"-->
</html>
