<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("admin_account")=""then
Response.write"对不起，您还没有登陆，无此权限！"
Response.end
end if
%>


<%
dim admin_account
admin_account=session("admin_account")
%>
<%
set rsc=server.createobject("adodb.recordset")
sql="select * from time_set where sorts='sheet'"
rsc.open sql,conn,1,3
%>

<%
if trim(request("select"))<>"" then
rsc("select")=trim(request("select"))
response.write"<script>alert('设置登录成绩状态成功！')</script>"
end if
%>

<%
if trim(request("enddate"))<>"" then
rsc("enddate")=trim(request("enddate"))
response.write"<script>alert('设置登录成绩结束日期成功！\n     "&rsc("enddate")&"')</script>"
end if
if trim(request("startdate"))<>"" then
rsc("startdate")=trim(request("startdate"))
response.write"<script>alert('设置登录成绩起始日期成功！\n     "&rsc("startdate")&"')</script>"
end if
%>

<%
rsc.update
%>

<script language="javascript">
	function checkcourse(form)
	{
		if (document.form.user_number.value=="")
		{
			alert("请输入学生学号！");
		}
		else if (document.form.score.value=="")
		{
			alert("请输入成绩！");
		}
		else
		{
			form.submit();
		}
		return false;
	}
</script>

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
.style12 {color: #006699;
	font-size: 12px;
}
-->
</style>
<style type="text/css">
<!--
.style14 {color: #FF6633;
	font-size: 12px;
}
.style15 {font-size: 11px}
-->
</style>
<style type="text/css">
<!--
.style16 {color: #FF6600}
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
                                <td><div align="center"><a href="sheet_set.asp"><img src="adminimages/scorequerry.gif" width="81" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="subject_set.asp"><img src="adminimages/examcourse.gif" width="81" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="sheet_add.asp"><img src="adminimages/scoreadd.gif" width="81" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="subject_add.asp"><img src="adminimages/courseadd1.gif" width="81" height="24" border="0"></a></div></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="86" background="adminimages/titlebk2.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="45"><div align="center"><img src="adminimages/scoremag.gif" width="523" height="45"></div></td>
                            </tr>
                        </table></td>
                      </tr>
					  <tr>
					  <td background="../user/userimages/titlebk.gif"><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <form name="form1" method="post" action="">
                            <td height="30" colspan="2">&nbsp;&nbsp;&nbsp;<span class="style12">登录成绩开始时间设定&nbsp;：</span>
                                <input name="startdate" type="text" class="style3" id="startdate" size="20" maxlength="20" value="<%=rsc("startdate")%>">
                              &nbsp;<span class="style3">必须按格式输入：2005-05-18，按回车确认</span></td>
                          </form>
                        </tr>
                        <tr>
                          <form name="form1" method="post" action="">
                            <td height="30" colspan="2">&nbsp;&nbsp;&nbsp;<span class="style12">登录成绩截止时间设定&nbsp;：</span>
                                <input name="enddate" type="text" class="style3" id="enddate" size="20" maxlength="20" value="<%=rsc("enddate")%>">
                              &nbsp;<span class="style3">必须按格式输入：2005-05-18，按回车确认</span></td>
                          </form>
                        </tr>
                        <tr>
                          <td height="30"><div align="center">
                              <input name="select" type="button" class="style8" id="select" onclick ="window.location.href='sheet_add.asp?select=yes'" value="开始登录成绩">
                          </div></td>
                          <td><div align="center">
                              <input name="select" type="button" class="style3" id="select" onclick ="window.location.href='sheet_add.asp?select=no'" value="强行中止登录">
                          </div></td>
                        </tr>
                        <tr>
                          <td height="25" colspan="2"><div align="center"></div></td>
                        </tr>
                        <tr>
                          <td colspan="2">&nbsp;</td>
                        </tr>
                      </table></td>
					  </tr>
                      <tr>
                        <td valign="top" background="../user/userimages/titlebk.gif"><div align="center">
                          <table width="90%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><div align="center">
                                <table width="200" border="1">
                                  <tr>
                                    <td><div align="center"><a href="sheet_addadmin.asp?academy=physics">物理系</a></div></td>
                                    <td><DIV align="center"><a href="sheet_addadmin.asp?academy=">中文系</a></DIV></td>
                                    <td><DIV align="center"><a href="sheet_addadmin.asp?academy=计算机系">计算机系</a></DIV></td>
                                  </tr>
                                </table>
                                
                              </div></td>
                            </tr>
                            <tr>
                              <td>&nbsp;</td>
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
