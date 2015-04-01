<!--#include file="conn.asp"-->
<!--#include file="Config.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("admin_account")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
dim admin_account,tutor_id
admin_account=session("admin_account")
tutor_id=session("tutor_id")
session("tutor_id")=tutor_id
%>

<html>
<head>
<script language="javascript">
window.status="欢迎访问南京大学物理系研究生管理信息系统！"
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>研究生信息管理系统</title>
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
                    <td width="60%" class="style3"><%=admin_account%>：</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td ><div align="center">                      </div>
                         <div align="center"></div></td>
                    <td height="30" class="style2">您已经<span class="style2">登录成功</span>,请</td>
                         <td>&nbsp;</td>
                   </tr>
                  <tr>
                    <td><div align="center"></div></td>
                    <td height="30" class="style2">选择您需要的服务!</td>
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
            <td height="35"><div align="center"><a href="http://www.nju.edu.cn/">南 京 大 学</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="http://physics.nju.edu.cn/">南 大 物 理 系</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="http://bbs.nju.edu.cn/">南 大 小 百 合</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="http://grawww.nju.edu.cn/">研 究 生 院</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="http://job.nju.edu.cn/">就 业 指 导 中 心</a></div></td>
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
                      <table width="80%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr>
                              <td height="25"><%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from info1 order by uploadtime desc"
	frst.open sql,myconn,1,1
	fcount = frst.recordcount
	if fcount > 0 then 	
		''显示参数
		dim tbbgcolor
		''分页参数
		dim maxperpage,pages,page
		maxperpage = 5
		frst.pagesize = maxperpage
		pages = frst.pagecount
		''页面参数设置
		page = Request.QueryString("page")
		if not isnumeric(page) then page = 1 else page = cint(page)
		if page < 1 then page = 1
		if page > pages then page = pages
		frst.absolutepage = page
		'显示内容
		'Set Isize=Server.CreateObject("WinImg.ImgSize")
		for i = 1 to maxperpage
			if frst.eof then exit for
			if i mod 2 = 0 then tbbgcolor = "" else tbbgcolor = "#ddeeff"
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
                                  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="">
                                    <tr class="style2">
                                      <td width="150"><div align="right">文件名称：</div></td>
                                      <td><a href="<%=fpath%>" target="_NEWwIN"><%=ftitle%>( 文件大小：<%=fsize%> bytes)</a> </td>
                                      <td align="right"></td>
                                    </tr>
                                    <tr class="style2">
                                      <td width="150"><div align="right">文件类型：</div></td>
                                      <td colspan="2"><%=ftype%></td>
                                    </tr>
                                    <tr class="style2">
                                      <td width="150"><div align="right">上传日期：</div></td>
                                      <td colspan="2"><%=fuploadtime%></td>
                                    </tr>
                                    <tr class="style2">
                                      <td width="150"><div align="right">说明：</div></td>
                                      <td colspan="2"><%=fdesc%></td>
                                    </tr>
                                    <tr bgcolor="">
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
                                    <tr class="style2">
                                      <td bgcolor="">还没有上传内容！</td>
                                    </tr>
                                  </table>
                                  <%
	  end if
	  frst.close
	  set frst = nothing
	  myconn.close
	  set myconn = nothing
	  %>
                                  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="">
                                    <tr class="style2">
                                      <td align="center"><%
		  If Page > 1 Then Response.Write ("<a href='?page=1'>首页</a><a href='?page="& Page - 1 &"'>上一页</a>")
		  If page < pages Then Response.Write ("&nbsp;<a href='?page="& Page + 1 &"'>下一页</a>&nbsp;<a href='?page="& Pages &"'>末页</a>")
		  %></td>
                                    </tr>
                                </table></td>
                            </tr>
                          </table></td>
                        </tr>
                        <tr>
                          <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr>
                              <td height="25" bgcolor=""><table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
                                  <form action="SaveUpload.asp" method="post" enctype="multipart/form-data" name="form1">
                                    <tr class="style2">
                                      <td width="33%" height="25" bgcolor=""><div align="left" class="style2">上传内容 </div></td>
                                      <td><%
		Response.Write "  允许上传的文件类型:<br> "
		Set Fs = Server.CreateObject("Scripting.FileSystemObject")
		For Each str In OKAr
		If Fs.FileExists(Server.MapPath("Images/"& Str &".gif")) Then
		Response.Write "<img src='Images/" & str & ".gif' alt='" & str & "文件'> "
		Else
		Response.Write "<img src='Images/X.gif' alt='" & str & "文件'> "
		End If
		Next
		Set Fs = Nothing
		Response.Write "<br>  允许上传的文件最大:"&Oksize / 1024&"KB"
		Response.Write "<br>  上传图片的大小: 100*120"
		%></td>
                                    </tr>
                                    <tr class="style2">
                                      <td width="33%"><div align="right">文件名称：</div></td>
                                      <td><input name="filetitle" type="text" class="style3" size="25">
                                          <select name="filetype" class="style3" id="filetype">
                                            <option value="学生照片">学生照片</option>
                                          </select>
                                      </td>
                                    </tr>
                                    <tr class="style2">
                                      <td valign="top"><div align="right">上传的文件：</div></td>
                                      <td><input name="filedata" type="file" class="TextBoxT" id="filedata" size="27"></td>
                                    </tr>
                                    <tr class="style2">
                                      <td valign="top"><div align="right">文件说明：<br>
                                      </div></td>
                                      <td><textarea name="filedesc" cols="36" rows="4" class="style3" id="filedesc"></textarea></td>
                                    </tr>
                                    <tr>
                                      <td>&nbsp;</td>
                                      <td><input name="Submit" type="submit" class="style2" value="上传内容"></td>
                                    </tr>
                                  </form>
                              </table></td>
                            </tr>
                          </table></td>
                        </tr>
                      </table>
                      <p>&nbsp;</p>
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
