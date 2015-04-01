<!--#include file="Config.asp"-->
<!--#include file="../admin/session.asp"-->
<script language="JavaScript" type="text/JavaScript">
var check=0
function checkall(form) { //v2.0
  if(check==0){
  for(var i=0;i<form.elements.length;i++)
  {
  var e=form.elements[i];
  e.checked=true;
  }
  check=1;
  chk.alt="全否";
  }else{
  for(var i=0;i<form.elements.length;i++)
  {
  var e=form.elements[i];
  e.checked=false;
  }
  check=0;
  chk.alt="全选";
  }
}
</script>


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
      <td height="40" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;您的当前位置：&gt;&gt; <span class="style3"><a href="index.asp">上传文件</a></span> -- <span class="style3">文件管理</span> -- <a href="Admin_main.asp">系统配置</a>&nbsp;&nbsp; <a href="../admin/admin_admin_logout.asp">［<span class="style2">注销</span>］</a></td>
      <td width="20" rowspan="2" background="../images/rightbk.jpg">&nbsp;</td>
    </tr>
    <tr>
      <td><div align="center">
          <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#0153A3">
            <tr class="text">
              <td height="10" align="center" bgcolor="#FFFFFF">&nbsp;</td>
            </tr>
            <tr class="text">
              <td height="50" align="center" bgcolor="#FFFFFF">后台管理-文件管理&nbsp;&nbsp;[<a href="Admin_main.asp">点此进入系统设置</a>]&nbsp;&nbsp;</td>
            </tr>
            <form name="del" action="del.asp" method="post">
              <tr>
                <td height="25" bgcolor="#FFFFFF">
	<%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from Fupload order by uploadtime desc"
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
		''显示内容
'Set Isize=Server.CreateObject("WinImg.ImgSize")
		for i = 1 to maxperpage
			if frst.eof then exit for
			if i mod 2 = 0 then tbbgcolor = "" else tbbgcolor = "#0066cc"
			fid = frst("Fupload_ID").Value
			ftitle = frst("FileTitle").Value
			fdesc = frst("FileDesc").Value
			ftype = frst("FileType").Value
			fpath = frst("FilePath").Value
			fsize = frst("Filesize").Value
			fuploadtime = frst("UploadTime").Value
'FileNameStr=Server.Mappath(fpath)
'Isize.GetImgSize Cstr(FileNameStr)
	%>
                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
                      <tr class="text">
                        <td width="150"><div align="right">文件名称：</div></td>
                        <td><a href="<%=fpath%>" target="_NEWwIN"><%=ftitle%>( 文件大小：<%=fsize%> bytes)</a> </td>
                        <td width="20%">删除
                            <input type="checkbox" name="DelID" value="<%=fid&"|"&fpath%>">
                        </td>
                      </tr>
                      <tr class="text">
                        <td width="150"><div align="right">文件类型：</div></td>
                        <td><%=ftype%></td>
                        <td><a href="Edit.asp?ID=<%=Fid%>">编辑</a></td>
                      </tr>
                      <tr class="text">
                        <td width="150"><div align="right">上传日期：</div></td>
                        <td><%=fuploadtime%></td>
                        <td><%
			Set Fs = Server.CreateObject("Scripting.FileSystemObject")
			If Fs.FileExists(server.mappath(fPath)) Then
			Response.Write("<img src=Images/isexists.gif")
			End If
			%>
                        </td>
                      </tr>
                      <tr class="text">
                        <td width="150"><div align="right">说明：</div></td>
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
                        <td bgcolor="#FFFFFF">还没有上传的内容！</td>
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
                        <td align="right" class="heading">&nbsp;
                            <%
		  If Page > 2 Then Response.Write ("<a href='?page=1'>首页</a><a href='?page="& Page - 1 &"'>上一页</a>")
		  If page < pages Then Response.Write ("&nbsp;<a href='?page="& Page + 1 &"'>下一页</a>&nbsp;<a href='?page="& Pages &"'>末页</a>")
		  %>
                        选中所有
                        <input name="chkall" type="checkbox" id="chkall" value="select" onclick=checkall(this.form)>
                        <input name="Submit" type="submit" class="style2" value="删除所选">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
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
