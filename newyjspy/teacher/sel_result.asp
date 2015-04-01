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
dim admin_account
admin_account=session("admin_account")
%>
<%
dim keywords,search_class,search_class1
keywords=trim(request("keywords"))
search_class=trim(request("search_class"))
search_class1=search_class
session("keywords")=keywords
session("search_class")=search_class
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
sql="select user_info.user_name,user_info.user_number,course.course_name,course.course_number,course_tutor,course_credit,course_term,course.course_info from user_info inner join (course_sel inner join course on course_sel.course_ID = course.course_ID) on user_info.user_ID = course_sel.user_ID where "&search_class&" like '%"&keywords&"%' and course_sel.selTime>'"&startdate&"' order by course.course_term,user_info.user_number desc"
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
<!--#include file="top1.asp"-->
<table width="800"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="187" height="3"></td>
    <td rowspan="2" valign="top"><div align="right">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="5"><div align="right"></div></td>
        </tr>
        <tr>
          <td valign="top"><div align="right">
            <table width="603"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="406" height="10">&nbsp;</td>
                <td width="406" background="../indeximages/midLinkTop.gif">&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td colspan="3"><div align="center">
                    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td height="54" background="adminimages/titlebk1.gif"><div align="center">
                            <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td><div align="center"><a href="sel_search.asp"><img src="adminimages/querryresult.gif" width="134" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="course_set.asp"><img src="adminimages/selcourse.gif" width="134" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="course_add.asp"><img src="adminimages/courseadd.gif" width="134" height="24" border="0"></a></div></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="86" background="adminimages/titlebk2.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="45"><div align="center"><img src="adminimages/selset.gif" width="523" height="45"></div></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr>
                        <td background="../user/userimages/titlebk.gif"><div align="center">
                            <table width="90%"  border="0" cellpadding="0" cellspacing="0">
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
                                      <tr>
                                        <td height="25" colspan="8"><div align="center"><span class="style2">&quot;</span><span class="style3"><%=rs("user_name")%></span><span class="style2">&quot;同学选课结果（学号 <%=rs("user_number")%>）</span></div></td>
                                      </tr>
                                      <tr>
                                        <td height="24"><div align="center"><span class="style10">课程</span></div></td>
                                        <td width="40" height="24"><div align="center" class="style10">学分</div></td>
                                        <td><div align="center" class="style10">授课教师</div></td>
                                      </tr>
                                      <%
	rs.move (Cint(NoncePage)-1)*24,1
	for i=1 to rs.pagesize
%>
                                      <tr>
                                        <td height="24"><div align="center" class="style10">
                                            <div align="center" class="style10"><%=rs("course_name")%></div>
                                        </div></td>
                                        <td width="40" height="24"><div align="center" class="style10"><%=rs("course_credit")%></div></td>
                                        <td><div align="center" class="style10"><%=rs("course_tutor")%></div></td>
                                      </tr>
                                      <%rs.movenext
if rs.eof then exit for
	next
rs.close
set rs=nothing
%>
                                      <tr>
                                        <td height="24" colspan="12"><div align="right"> <span class="style10">
                                            <input type="hidden" name="page" value="<%=NoncePage%>">
                                            <%
if NoncePage>1 then
	response.write "|<a href=sel_result.asp?page=1&keywords="&keywords&"&search_class="&search_class1&">首 页</a>| |<a href=sel_result.asp?page="&NoncePage-1&"&keywords="&keywords&"&search_class="&search_class1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=sel_result.asp?page="&NoncePage+1&"&keywords="&keywords&"&search_class="&search_class1&">下一页</a>| |<a href=sel_result.asp?page="&NumPage&"&keywords="&keywords&"&search_class="&search_class1&">尾 页</a>|"
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
                                      <tr>
                                        <td height="24" colspan="7"><div align="right"><span class="style2">&quot;</span><span class="style3"><%=rs("course_name")%></span><span class="style2">&quot;课程选课结果&nbsp;</span><span class="style2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><span class="style2">&nbsp;&nbsp;&nbsp;授课老师&nbsp;&nbsp;&nbsp;</span><span class="style2">&nbsp;&nbsp;&nbsp;<%=rs("course_tutor")%>&nbsp;&nbsp;&nbsp;</span><span class="style2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></div></td>
                                      </tr>
                                      <tr>
                                        <td height="24"><div align="center"><span class="style10">姓名</span></div></td>
                                        <td><div align="center"class="style10">学号</div></td>
                                        <td height="24"><div align="center" class="style10">学分</div></td>
                                        <td><div align="center" class="style10"></div></td>
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
                                        <td>&nbsp;</td>
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
	response.write "|<a href=sel_result.asp?page=1&keywords="&keywords&"&search_class="&search_class1&">首 页</a>| |<a href=sel_result.asp?page="&NoncePage-1&"&keywords="&keywords&"&search_class="&search_class1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=sel_result.asp?page="&NoncePage+1&"&keywords="&keywords&"&search_class="&search_class1&">下一页</a>| |<a href=sel_result.asp?page="&NumPage&"&keywords="&keywords&"&search_class="&search_class1&">尾 页</a>|"
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
                                <td height="30"><div align="center"><a href="sel_print.asp"><img src="../user/userimages/print1.gif" width="51" height="23" border="0" align="absmiddle"></a></div></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="40" background="adminimages/titlebk3.gif">&nbsp;</td>
                      </tr>
                    </table>
                </div></td>
              </tr>
            </table>
          </div></td>
        </tr>
        <tr>
          <td height="15" valign="top"><div align="right">
          </div></td>
        </tr>
        <tr>
          <td valign="top"><div align="right">
            <!--#include file="server.asp"-->
		  </div></td>
        </tr>
      </table>
    </div></td>
  </tr>
  <tr>
    <td valign="top" background="../indeximages/loginbk.gif"><div align="center">
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="47" background="adminimages/adminlogin.gif">&nbsp;</td>
          </tr>
          <tr>
            <td><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="100"><div align="center">
                    <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" background="../indeximages/loginbk.gif">
                      <tr>
                        <td width="20%"><div align="center"> </div>
                            <div align="center"></div></td>
                        <td width="60%" class="style3"><%=admin_account%>：</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr>
                        <td ><div align="center"> </div>
                            <div align="center"></div></td>
                        <td height="30" class="style2">您已经登录成功,可</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr>
                        <td><div align="center"></div></td>
                        <td height="30" class="style2">以正常维护网站!</td>
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
                <td height="6" background="../indeximages/loginbk.gif"><div align="center"><img src="../indeximages/loginbar.gif" width="129" height="2"></div></td>
              </tr>
              <tr>
                <td height="60" valign="center" background="../indeximages/loginbk.gif"><div align="center"><a href="admin_logout.asp"><img src="../includeimages/logout.gif" width="60" height="24" border="0"></a></div></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td height="77" background="../indeximages/links.gif">&nbsp;</td>
          </tr>
          <tr>
            <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><div align="center">
                    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td height="25"><div align="center"></div></td>
                      </tr>
                      <script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//--></script>
                <tr>
                  <td height="50"><div align="center">
                      <form name="links">
                        <select name="links" class="style2" onChange="window.open(this.value)">
                          <option value="javascript:void(null);" selected>----国外大学----</option>
                          <option value="http://www.harvard.edu/">哈佛大学</option>
                          <option value="http://www.cam.ac.uk/">剑桥大学</option>
                          <option value="http://www.ox.ac.uk/">牛津大学</option>
                          <option value="http://www.stanford.edu/">斯坦福大学</option>
                          <option value="http://www.yale.edu/">耶鲁大学</option>
                        </select>
                            </form>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="50"><div align="center">
                            <form name="links">
                              <select name="links" class="style2" onChange="window.open(this.value)">
                                <option value="javascript:void(null);" selected>---实验室链接---</option>
                                <option value="http://biophy.nju.edu.cn ">生物物理实验室</option>
                                <option value="http://pld.nju.edu.cn ">PLD实验室</option>
                                <option value="http://x.nju.edu.cn/">邢定钰小组</option>
                              </select>
                            </form>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="50"><div align="center">
                            <form name="links">
                              <select name="links" class="style2" onChange="window.open(this.value)">
                                <option value="javascript:void(null);" selected>----校外链接----</option>
                                <option value="http://www.njbys.com/">南京毕业生就业网</option>
                                <option value="http://www.jsbys.com.cn/index.aspx">江苏毕业生就业网</option>
                                <option value="http://www.firstjob.com.cn/">上海毕业生就业网</option>
                              </select>
                            </form>
                        </div></td>
                      </tr>
                    </table>
                </div></td>
              </tr>
              <tr>
                <td height="35"><div align="center"><a href="../links.asp" class="style3">&gt;&gt;&gt; More</a></div></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table>
      </div></td>
  </tr>
  <tr>
    <td height="34" background="../indeximages/loginbottom.gif">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<!--#include file="bottom1.asp"-->
</html>
