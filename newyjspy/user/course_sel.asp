<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("user_account")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<!--#include file="regfirst.asp"--> 
<%
	dim today
	today=Date 
	today=Year(today) & "-" & Right("0" & Month(today),2) & "-" & Right("0" & Day(today),2)
%>

<%
dim enddate,sel,startdate
set rscs=server.createobject("adodb.recordset")
sql="select * from course_set"
rscs.open sql,conn,1,1
enddate = rscs("enddate")
startdate=rscs("startdate")
sel = rscs("select")
%>

<%
set rsu=server.createobject("adodb.recordset")
sql="select course_sel.seltime,course_sel.course_selID,course.course_number,course.course_name,course.course_tutor,course.course_credit,course.course_term,course.course_time from user_info inner join (course_sel inner join course on course_sel.course_ID=course.course_ID) on user_info.user_ID=course_sel.user_ID where course_sel.selTime >'"&startdate&"' and user_account='"&session("user_account")&"'"
rsu.open sql,conn,1,1
%>
<%
set rsc=server.createobject("adodb.recordset")
sql="select * from course where course_term = '1' order by course_term "
rsc.open sql,conn,1,1
%>



<%
if today < startdate then
Response.write "<script>alert('本学期选课尚未开始！');history.go(-1);</script>"
end if
%>
<%
if today > enddate or sel = "no" then
Response.write "<script>alert('本学期选课已结束或已中止！');</script>"
%>
<html>
<head>
<script language="javascript">
<!--

window.status="欢迎访问南京大学物理系研究生管理信息系统！"
//-->
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
.style15 {
	color: #FF6600;
	font-size: 15px;
}
-->
</style>
<!--#include file="top1.asp"-->
<table width="800"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="187" height="10"></td>
    <td rowspan="2" valign="top"><div align="right">
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="22"><div align="right"></div></td>
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
                            <td height="54" background="userimages/titlebk1.gif"><div align="center"><img src="userimages/courseSel.gif" width="523" height="45"></div></td>
                          </tr>
                          <tr>
                            <td height="53" valign="bottom" background="userimages/titlebk2.gif"><div align="center"><a href="course_sheet.html" target="_blank"><font  class="style15" >查 看 课 程 表</font></a> </div></td>
                          </tr>
                          <tr>
                            <td height="25" valign="bottom" background="userimages/titlebk.gif"><div align="center"></div></td>
                          </tr>
                          <tr>
                            <td background="userimages/titlebk.gif"><div align="center">
                                <table width="90%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td valign="top">
<%
if (rsu.bof and rsu.eof) then
	response.write "<table width='100%' height='30' border='1'cellpadding='0' cellspacing='0'bordercolor='#000000' class='thin'><tr><td colspan=13><marquee scrolldelay=120 behavior=alternate><font class='style3' >"&rs("user_account")&"  您还没有选择课程！</font></marquee></td></tr></table>"
else
%>
                                        <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                          <tr>
                                            <td height="24" colspan="6"><div align="center" class="style3">已 选 课 程</div></td>
                                          </tr>
                                          <tr>
                                            <td height="24"><div align="center" class="style10">课程名称</div></td>
                                            <td height="24"><div align="center" class="style10">主讲老师</div></td>
                                            <td height="24"><div align="center" class="style10">学分</div></td>
                                            <td height="24"><div align="center" class="style10">上课时间</div></td>
                                          </tr>
                                          <%
	for i=1 to rsu.recordcount
%>
                                          <tr>
                                            <td height="24"><div align="left" class="style10">&nbsp;&nbsp;&nbsp;&nbsp;<%=rsu("course_name")%></div></td>
                                            <td height="24"><div align="center" class="style10"><%=rsu("course_tutor")%></div></td>
                                            <td height="24"><div align="center" class="style10"><%=rsu("course_credit")%></div></td>
                                            <td height="24"><div align="left" class="style10">&nbsp;<%=rsu("course_time")%></div></td>
                                          </tr>
                                          <%rsu.movenext
	next

rsu.close
set rsu=nothing
%>
                                        </table>
                                        <%
end if
%>
                                      </td>
                                  </tr>
                                  <tr>
                                    <td></td>
                                  </tr>
                                </table>
                            </div></td>
                          </tr>
                          <tr>
                            <td height="34" background="userimages/titlebk3.gif">&nbsp;</td>
                          </tr>
                        </table>
                    </div></td>
                  </tr>
                </table>
            </div></td>
          </tr>
          <tr>
            <td height="15" valign="top"><div align="right"> </div></td>
          </tr>
          <tr>
            <td valign="top"><div align="right">
            </div></td>
          </tr>
        </table>
    </div></td>
  </tr>
  <tr>
    <td valign="top" background="../indeximages/loginbk.gif"><div align="center"> </div>
        <div align="center">
          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="47" background="../indeximages/stulogin.gif">&nbsp;</td>
            </tr>
            <tr>
              <td><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td height="100"><div align="center">
                        <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" background="../indeximages/loginbk.gif">
                          <tr>
                            <td width="20%"><div align="center"> </div>
                                <div align="center"></div></td>
                            <td width="60%" class="style3"><%=rs("user_account")%>：</td>
                            <td>&nbsp;</td>
                          </tr>
                          <tr>
                            <td ><div align="center"> </div>
                                <div align="center"></div></td>
                            <td height="30" class="style2">您已经登录成功,请</td>
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
                    <td height="6" background="../indeximages/loginbk.gif"><div align="center"><img src="../indeximages/loginbar.gif" width="129" height="2"></div></td>
                  </tr>
                  <tr>
                    <td height="60" valign="center" background="../indeximages/loginbk.gif"><div align="center"><a href="user_logout.asp"><img src="../includeimages/logout.gif" width="60" height="24" border="0"></a></div></td>
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

<%
else
%>
<html>
<head>
<script language="javascript">
<!--

window.status="欢迎访问南京大学物理系研究生管理信息系统！"
//-->
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
.style15 {
	color: #FF6600;
	font-size: 15px;
}
-->
</style>
<!--#include file="top1.asp"-->
<table width="800"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="187" height="10"></td>
    <td rowspan="2" valign="top"><div align="right">
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="22"><div align="right"></div></td>
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
                            <td height="54" background="userimages/titlebk1.gif"><div align="center"><img src="userimages/courseSel.gif" width="523" height="45"></div></td>
                          </tr>
                          <tr>
                            <td height="53" valign="bottom" background="userimages/titlebk2.gif"><div align="center"><a href="course_sheet.html" target="_blank"><font  class="style15" >查 看 课 程 表</font></a> </div></td>
                          </tr>
                          <tr>
                            <td height="25" valign="bottom" background="userimages/titlebk.gif"><div align="center">（本次选课时间：<%=startdate%> --- <%=enddate%>） </div></td>
                          </tr>
                          <tr>
                            <td background="userimages/titlebk.gif"><div align="center">
                                <table width="90%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td valign="top"><%
if (rsu.bof and rsu.eof) then
	response.write "<table width='100%' height='30' border='1'cellpadding='0' cellspacing='0'bordercolor='#000000' class='thin'><tr><td colspan=13><marquee scrolldelay=120 behavior=alternate><font class='style3' >"&rs("user_account")&"  您还没有选择课程！</font></marquee></td></tr></table>"
else
%>
                                        <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                          <tr>
                                            <td height="24" colspan="6"><div align="center" class="style3">已 选 课 程</div></td>
                                          </tr>
                                          <tr>
                                            <td height="24"><div align="center" class="style10">课程名称</div></td>
                                            <td height="24"><div align="center" class="style10">主讲老师</div></td>
                                            <td height="24"><div align="center" class="style10">学分</div></td>
                                            <td height="24"><div align="center" class="style10">上课时间</div></td>
                                            <td height="24"><div align="center" class="style10">删除</div></td>
                                          </tr>
                                          <%
	for i=1 to rsu.recordcount
%>
                                          <tr>
                                            <td height="24"><div align="left" class="style10">&nbsp;&nbsp;&nbsp;&nbsp;<%=rsu("course_name")%></div></td>
                                            <td height="24"><div align="center" class="style10"><%=rsu("course_tutor")%></div></td>
                                            <td height="24"><div align="center" class="style10"><%=rsu("course_credit")%></div></td>
                                            <td height="24"><div align="left" class="style10">&nbsp;<%=rsu("course_time")%></div></td>
                                            <td height="24"><div align="center"><a href=course_seldel.asp?course_selID=<%=rsu("course_selID")%>><font color="#ff6633">删除</font></a></div></td>
                                          </tr>
                                          <%rsu.movenext
	next

rsu.close
set rsu=nothing
%>
                                        </table>
                                        <%
end if
%>
                                        <form name="form1" method="post" action="course_sel_add.asp">
                                          <%
if (rsc.bof and rsc.eof) then
	response.write "<table width='100%' height='30' border='1'cellpadding='0' cellspacing='0' bordercolor='#000000' class='thin'><tr><td colspan=13><marquee scrolldelay=120 behavior=alternate><font class='style3'>尚未设置选择课程！</font></marquee></td></tr></table>"
else
%>
                                          <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                            <tr>
                                              <td height="24" colspan="6"><div align="center" class="style3">待 选 课 程</div></td>
                                            </tr>
                                            <tr>
                                              <td width="30"><div align="center" class="style10">状态</div></td>
                                              <td height="24"><div align="center" class="style10">课程名称</div></td>
                                              <td height="24"><div align="center" class="style10">主讲老师</div></td>
                                              <td height="24"><div align="center" class="style10">学分</div></td>
                                              <td height="24"><div align="center" class="style10">上课时间</div></td>
                                            </tr>
                                            <%
	for i=1 to rsc.recordcount
%>
                                            <tr>
                                              <td><div align="center" class="style10">
                                                  <input type="radio" name="course_ID" value="<%=rsc("course_ID")%>">
                                              </div></td>
                                              <td height="24"><div align="left" class="style10">&nbsp;&nbsp;&nbsp;&nbsp;<%=rsc("course_name")%></div></td>
                                              <td height="24"><div align="center" class="style10"><%=rsc("course_tutor")%></div></td>
                                              <td height="24"><div align="center" class="style10"><%=rsc("course_credit")%></div></td>
                                              <td height="24"><div align="left" class="style10">&nbsp;<%=rsc("course_time")%></div></td>
                                            </tr>
                                            <%rsc.movenext
	next
rsc.close
set rsc=nothing
%>
                                            <tr>
                                              <td height="35" colspan="5"><div align="center"><img src="userimages/courseAdd.gif" width="70" height="25" align="absmiddle" style="cursor:hand; " onMouseDown="submit()"> </div></td>
                                            </tr>
                                          </table>
                                          <%
end if
%>
                                      </form></td>
                                  </tr>
                                  <tr>
                                    <td></td>
                                  </tr>
                                </table>
                            </div></td>
                          </tr>
                          <tr>
                            <td height="34" background="userimages/titlebk3.gif">&nbsp;</td>
                          </tr>
                        </table>
                    </div></td>
                  </tr>
                </table>
            </div></td>
          </tr>
          <tr>
            <td height="15" valign="top"><div align="right"> </div></td>
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
    <td valign="top" background="../indeximages/loginbk.gif"><div align="center"> </div>
        <div align="center">
          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="47" background="../indeximages/stulogin.gif">&nbsp;</td>
            </tr>
            <tr>
              <td><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td height="100"><div align="center">
                        <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" background="../indeximages/loginbk.gif">
                          <tr>
                            <td width="20%"><div align="center"> </div>
                                <div align="center"></div></td>
                            <td width="60%" class="style3"><%=rs("user_account")%>：</td>
                            <td>&nbsp;</td>
                          </tr>
                          <tr>
                            <td ><div align="center"> </div>
                                <div align="center"></div></td>
                            <td height="30" class="style2">您已经登录成功,请</td>
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
                    <td height="6" background="../indeximages/loginbk.gif"><div align="center"><img src="../indeximages/loginbar.gif" width="129" height="2"></div></td>
                  </tr>
                  <tr>
                    <td height="60" valign="center" background="../indeximages/loginbk.gif"><div align="center"><a href="user_logout.asp"><img src="../includeimages/logout.gif" width="60" height="24" border="0"></a></div></td>
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
<%
end if
%>
<%
rscs.close
set rscs=nothing
%>
