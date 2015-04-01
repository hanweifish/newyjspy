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

<%
dim user_account
user_account=session("user_account")
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
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

<script language="javascript">
	function checkuser(form)
	{
		if (document.form.user_pwd.value=="")
		{
			alert("请输入密码！!");
		}
		else if (document.form.user_pwd1.value=="")
		{
			alert("请输入确认密码！!");
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
.style13 {color: #FF0000}
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
                        <td height="54" background="userimages/titlebk1.gif"><div align="center"><img src="userimages/stuinfo.gif"></div></td>
                      </tr>
                      <tr>
                        <td height="53" background="userimages/titlebk2.gif">&nbsp;</td>
                      </tr>
                      <tr>
                        <td background="userimages/titlebk.gif"><div align="center">
                            <form name="form" method="post" action="info_mod1.asp">
                              <div align="center">
                                <table width="80%" height="100%" border="0" cellpadding="0" cellspacing="0" class="thin">
                                  <tr>
                                    <td width="32%" height="24"><div align="right" class="style10">用户名：&nbsp;&nbsp;</div></td>
                                    <td width="68%" class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_account")%></td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">姓名：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_name")%> </td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">学号：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_number")%></td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">密码：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input name="user_pwd" type="password" class="style3" value="<%=rs("user_pwd")%>" size="24" >
                                        <span class="style13">*</span> </td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">密码确认：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input name="user_pwd1" type="password" class="style3" value="<%=rs("user_pwd")%>"  size="24" >
                                        <span class="style13">*</span> </td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">专业：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("user_major")%>                                        </td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">年级：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <select name="user_grade" class="style3" id="user_grade">
                                          <%
if rs("user_grade") = "研一" then
%>
                                          <option value="研一" selected>研一</option>
                                          <option value="研二">研二</option>
                                          <option value="研三">研三</option>
                                          <%
elseif rs("user_grade") = "研二" then
%>
                                          <option value="研一">研一</option>
                                          <option value="研二" selected>研二</option>
                                          <option value="研三">研三</option>
                                          <%
elseif rs("user_grade") = "研三" then
%>
                                          <option value="研一">研一</option>
                                          <option value="研二">研二</option>
                                          <option value="研三" selected>研三</option>
                                          <%
end if
%>
                                        </select>
                                        <span class="style13">*</span></td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">E-mail：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input name="user_mail" type="text" class="style3" value="<%=rs("user_mail")%>" size="24" >
                                        <span class="style13">*</span> </td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">BBS帐号：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input name="user_bbs" type="text" class="style3" id="user_bbs" value="<%=rs("user_bbs")%>" size="24" >                                    </td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">联系方式：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input name="user_roomphone" type="text" class="style3" value="<%=rs("user_roomphone")%>" size="24" >
                                        <span class="style13">*</span> （形式：025-83594521） </td>
                                  </tr>
                                  
                                  <tr>
                                    <td height="24"><div align="right" class="style10">家庭电话：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input name="user_homephone" type="text" class="style3" value="<%=rs("user_homephone")%>" size="24" >
                                        <span class="style13">*</span>&nbsp;&nbsp;（形式：025-83594521） </td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">家庭地址：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input name="user_address" type="text" class="style3" value="<%=rs("user_address")%>" size="24" >
                                        <span class="style13">*</span> </td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">邮编：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input name="user_code" type="text" class="style3" value="<%=rs("user_code")%>" size="24" >
                                        <span class="style13">*</span> </td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">性别：&nbsp;&nbsp;</div></td>
                                    
                                    <td class="style10">&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("user_sex")%></td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">生日：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("user_Csrq")%></td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">备注：&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <textarea name="user_info" cols="45" rows="10"><%=HTMLEncode(rs("user_info"))%></textarea>                                    </td>
                                  </tr>
                                  <tr>
                                    <td height="35" colspan="2"><div align="center"><img src="userimages/editSub.gif" width="70" height="25" align="absmiddle" style="cursor:hand; " onClick="checkuser(form);"> </div></td>
                                  </tr>
                                </table>
                              </div>
                            </form>
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
      </div>      <div align="center">
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
