<!--#include file="Config.asp"-->
<!--#include file="../admin/session.asp"-->
<%
Dim ID,Msg
	ID = Request.QueryString("ID")
	If Request.QueryString("Action") = "Save" Then SaveData ID
	Sub SaveData(ID)
	If ID < 1 Then 
	Response.Write("参数错误")
	Response.End()
	End If
	myConn.execute("update Fupload set FILETITLE='"&Request.Form("fname")&"',FILEDESC='"&Request.Form("fdesc")&"',FILETYPE='"&Request.Form("ftype")&"',FILEPATH='"&Request.Form("fpath")&"',FILESIZE='"&Request.Form("fsize")&"' where Fupload_ID="&ID)
	Msg = "成功修改了文件数据信息"
	End Sub

If msg <> "" Then
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_List.asp'>"&Msg&"<br>本页将在3秒内返回<BR>如果你的浏览器没有反应，请<a href=Admin_List.asp>点击此处返回</a>")
Response.End()
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>上传管理</title>
<link href="../style.css" rel="stylesheet" type="text/css">
</head>

<body>
<div align="center">
  <!--#include file="top1.asp" -->
  <table width="840" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="20" rowspan="2" background="../images/leftbk.jpg">&nbsp;</td>
      <td height="40" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;您的当前位置：&gt;&gt; <span class="style3"><a href="Admin_main.asp">系统配置</a></span> &gt;&gt; <span class="style3">编辑文件属性</span>&nbsp;&nbsp; <a href="../admin/admin_logout.asp">［<span class="style2">注销</span>］</a></td>
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
                        <td width="150" bgcolor="#FFFFFF"><div align="right">文件名称：</div></td>
                        <td bgcolor="#FFFFFF"><input type="text" name="fname" class="TextBoxT" value="<%=ftitle%>">
                        </td>
                        <td width="30%" rowspan="3" bgcolor="#FFFFFF">&nbsp;</td>
                      </tr>
                      <tr class="text">
                        <td width="150" bgcolor="#FFFFFF"><div align="right">文件类型：</div></td>
                        <td bgcolor="#FFFFFF"><select name="ftype" class="TextBoxT" id="filetype">
                          <option value="文    件" selected>文 件</option>
                          <option value="图    片">图 片</option>
                          <option value="媒    体">媒 体</option>
                        </select></td>
                      </tr>
                      <tr class="text">
                        <td width="150" bgcolor="#FFFFFF"><div align="right">文件路径：</div></td>
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
                        <td width="150" bgcolor="#FFFFFF"><div align="right">说明：</div></td>
                        <td colspan="2" bgcolor="#FFFFFF"><input type="text" name="fdesc" class="TextBoxT" value="<%=fdesc%>"></td>
                      </tr>
                      <tr class="text">
                        <td height="1" align="right" bgcolor="#FFFFFF">文件大小：</td>
                        <td height="1" colspan="2" bgcolor="#FFFFFF"><input type="text" name="fsize" class="TextBoxT" value="<%=fsize%>">
                        bytes</td>
                      </tr>
                      <tr class="text">
                        <td height="1" align="right" bgcolor="#FFFFFF">&nbsp;</td>
                        <td height="1" colspan="2" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="修改">
&nbsp;
                        <input type="button" name="Submit2" value="返回" onclick='javascript:window.location="Admin_List.asp"'></td>
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
                        <td bgcolor="#FFFFFF">没有对应的数据！</td>
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