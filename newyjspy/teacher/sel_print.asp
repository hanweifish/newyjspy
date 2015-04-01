<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("admin_account")="" or session("user_group")="" then
Response.write"对不起，您还没有登陆，无此权限！"
Response.end
end if
%>

<%
dim enddate,startdate
set rsc=server.createobject("adodb.recordset")
sql="select * from course_set"
rsc.open sql,conn,1,3
startdate=rsc("startdate")
%>
<%
dim keywords,search_class
keywords=session("keywords")
search_class=session("search_class")
%>
<%
select case search_class
case "学生姓名"
search_class="user_info.user_name"
case "学生学号"
search_class="user_info.user_number"
case "课程"
search_class="course.course_name"
case Else
response.redirect "sel_search.asp"
End select
%>
<%
set rs=server.createobject("adodb.recordset")
sql="select user_info.user_name,user_info.user_number,course.course_name,course.course_number,course_tutor,course_credit,course_term,course.course_info from user_info inner join (course_sel inner join course on course_sel.course_ID = course.course_ID) on user_info.user_ID = course_sel.user_ID where "&search_class&" like '%"&keywords&"%'  and course_sel.selTime>'"&startdate&"' order by course.course_term,user_info.user_number desc"
rs.open sql,conn,1,1
%>
<%
function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
    HTMLEncode = fString
end if
end function
%>
<html>
<head>
<script language="javascript">
window.status="欢迎访问南京大学物理系研究生管理信息系统！"
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>研究生信息管理系统</title>
<link href="../style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style10 {font-size: 12px;
	color: #004080;
}
-->
</style>
<style type="text/css">
<!--
.style12 {color: #006699; font-size: 13px;}
-->
</style>
<style type="text/css">
<!--
.style13 {color: #006699;
	font-size: 12px;
}
.style14 {color: #FF6633;
	font-size: 12px;
}
.style15 {font-size: 11px}
-->
</style>
</head>
<div align="center">
  <table width="700"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="90%"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="60">&nbsp;</td>
        </tr>
        <tr>
          <td><%if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=24
			NumPage=rs.Pagecount
		if request("page")=empty then 
			NoncePage=1
		else
		if Cint(request("page"))<1 then
			NoncePage=1
		else
			NoncePage=request("page")
		end if
	if Cint(Trim(request("page")))>Cint(NumPage) then NoncePage=NumPage
	end if
else
	NumRecord=0
	NumPage=0
	NoncePage=0
end if
%>
              <%
if (rs.bof and rs.eof) then
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>暂时没有选课记录!!!</font></marquee></td></tr>"
else
if search_class="user_info.user_name" or search_class="user_info.user_number" then
%>
              <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                <tr bgcolor="#EEF4FF">
                  <td height="25" colspan="9"><div align="right"><span class="style2">&quot;</span><span class="style3"><%=rs("user_name")%></span><span class="style2">&quot;同学选课结果（学号 <%=rs("user_number")%>）<span class="style10">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;授课教师&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("course_tutor")%>&nbsp;&nbsp;&nbsp;&nbsp;</span></span></div></td>
                </tr>
                <tr>
                  <td width="207" height="24"><div align="center"><span class="style10">课 程</span></div></td>
                  <td width="71" height="24"><div align="center" class="style10">学分</div></td>
                  <td width="182"><div align="center" class="style10"></div></td>
                  <td width="160"><div align="center" class="style10">备注信息</div></td>
                </tr>
                <%
	rs.move (Cint(NoncePage)-1)*24,1
	for i=1 to rs.pagesize
%>
                <tr>
                  <td height="24"><div align="center" class="style10">
                      <div align="center" class="style10"><%=rs("course_name")%></div>
                  </div></td>
                  <td width="71" height="24"><div align="center" class="style10"><%=rs("course_credit")%></div></td>
                  <td><div align="center" class="style10"></div></td>
                  <td><div align="center" class="style10"><%=HTMLEncode(rs("course_info"))%></div></td>
                </tr>
                <%rs.movenext
if rs.eof then exit for
	next
rs.close
set rs=nothing
%>
                <tr>
                  <td height="24" colspan="13"><div align="right"> <span class="style10">
                      <input type="hidden" name="page" value="<%=NoncePage%>">
                      <%
if NoncePage>1 then
	response.write "|<a href=sel_result.asp?page=1>首 页</a>| |<a href=sel_result.asp?page="&NoncePage-1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=sel_result.asp?page="&NoncePage+1&">下一页</a>| |<a href=sel_result.asp?page="&NumPage&">尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
&nbsp;页次：<font color="#0033CC"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC"><%=NumRecord%></font>条记录</span>&nbsp; </div></td>
                </tr>
              </table>
              <%
else
%>
              <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                <tr bgcolor="#EEF4FF">
                  <td height="24" colspan="7"><div align="right"><span class="style2">&quot;</span><span class="style3"><%=rs("course_name")%></span><span class="style2">&quot;课程选课结果<span class="style10">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;授课教师&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("course_tutor")%>&nbsp;&nbsp;&nbsp;&nbsp;</span></span></div></td>
                </tr>
                <tr>
                  <td width="19%" height="24"><div align="center"><span class="style10">姓名</span></div></td>
                  <td width="27%"><div align="center"class="style10">学 号 </div></td>
                  <td width="27%" height="24"><div align="center" class="style10">学分</div></td>
                  <td width="27%"><div align="center" class="style10"></div></td>
                </tr>
                <%
	rs.move (Cint(NoncePage)-1)*24,1
	for i=1 to rs.pagesize
%>
                <tr>
                  <td height="24"><div align="center" class="style10">
                      <div align="center" class="style10"><%=rs("user_name")%></div>
                  </div></td>
                  <td height="24"><div align="center"class="style10"><%=rs("user_number")%></div></td>
                  <td height="24"><div align="center" class="style10"><%=rs("course_credit")%></div></td>
                  <td><div align="center" class="style10"></div></td>
                </tr>
                <%rs.movenext
if rs.eof then exit for
	next
rs.close
set rs=nothing
%>
                <tr>
                  <td height="24" colspan="7"><div align="right"> <span class="style10">
                      <input type="hidden" name="page" value="<%=NoncePage%>">
                      <%
if NoncePage>1 then
	response.write "|<a href=sel_result.asp?page=1>首 页</a>| |<a href=sel_result.asp?page="&NoncePage-1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=sel_result.asp?page="&NoncePage+1&">下一页</a>| |<a href=sel_result.asp?page="&NumPage&">尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
&nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录</span>&nbsp; </div></td>
                </tr>
              </table>
              <%
end if
end if	
%>
          </td>
        </tr>
        <tr>
          <td height="30"><div align="center"><a href="sel_print.asp"><img src="../user/userimages/print1.gif" width="51" height="23" border="0" onClick="javascript:window.print()"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="sel_search.asp"><img src="../user/userimages/return.gif" width="49" height="23" border="0" align="absmiddle"></a></div></td>
        </tr>
      </table></td>
    </tr>
  </table>
</div>
</body>
</html>
