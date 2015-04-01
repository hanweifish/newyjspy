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
dim admin_account
admin_account=session("admin_account")
%>

<script language="javascript">
	function checkform(form)
	{
		if (document.form.job_title.value=="")
		{
			alert("请输入通知标题！");
		}
		else if (document.form.job_content.value=="")
		{
			alert("请输入通知内容！");
		}
		else if (document.form.job_type.value=="----信息类型----")
		{
			alert("请选择信息类型！");
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
<!--#include file="top1.asp"-->
<table width="800"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="187" height="12"></td>
    <td></td>
  </tr>
  <tr>
    <td rowspan="3" valign="top"><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="47" background="adminimages/adminlogin.gif">&nbsp;</td>
      </tr>
      <tr>
        <td><div align="center">
          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="100"><div align="center">
                <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" background="../indeximages/loginbk.gif">
                  <tr>
                    <td width="20%"><div align="center">                      </div>
                         <div align="center"></div></td>
                    <td width="60%" class="style3"><%=admin_account%>：</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td ><div align="center">                      </div>
                         <div align="center"></div></td>
                    <td height="30" class="style2">您已经<span class="style2">登录成功</span>,可</td>
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
          </table>
        </div></td>
      </tr>
      <tr>
        <td height="77" background="../indeximages/links.gif">&nbsp;</td>
      </tr>
      <tr>
        <td background="../indeximages/loginbk.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
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
        <td height="34" background="../indeximages/loginbottom.gif">&nbsp;</td>
      </tr>
    </table></td>
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
                          <td><div align="center"></div>                            <div align="center"><a href="job_add.asp"><img src="adminimages/jobAdd.gif" width="240" height="24" border="0"></a></div>                            <div align="center"></div></td>
                          </tr>
                      </table>
                    </div></td>
                  </tr>
                  <tr>
                    <td height="86" background="adminimages/titlebk2.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td height="45"><div align="center"><img src="../indeximages/job.gif" width="523" height="45"></div></td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td background="../user/userimages/titlebk.gif"><div align="center">
                      <table width="90%"  border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td><form name="form" method="post" action="../admin/job_add1.asp">
                              <div align="center">
                                <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                  <tr>
                                    <td height="24"><div align="right" class="style10">信息类型：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp; <select name="job_type" class="style10" id="job_type">
                                      <option value="----信息类型----" selected>----信息类型----</option>
                                      <option value="invite">招聘信息</option>
                                      <option value="policy">就业政策</option>
                                    </select></td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">信息标题：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input name="job_title" type="text" class="style2" size="50"></td>
                                  </tr>
                                  <tr>
                                    <td ><div align="right" class="style10">信息内容：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <textarea name="job_content" cols="50" rows="15" class="style2"></textarea></td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">信息备注：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <textarea name="job_info" cols="50" rows="10" class="style2">无</textarea>
                                    </td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="center" class="style2">发布人：</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp; <input name="job_author" type="text" class="style2" id="job_author"></td>
                                  </tr>
                                  <tr>
                                    <td height="35" colspan="2"><div align="center"><img src="../user/userimages/add.gif" width="51" height="23" style="cursor:hand" onmousedown="checkform(form)"> </div></td>
                                  </tr>
                                </table>
                              </div>
                          </form></td>
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
