<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("admin_account")="" or session("user_group")="subadmin" then
Response.write"对不起，您还没有登陆或者无此权限！"
Response.end
end if
%>

<%
dim admin_account,user_id,NoncePage
NoncePage=trim(request("NoncePage"))
user_id=trim(request("user_ID"))
session("user_ID")=user_id
admin_account=session("admin_account")
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_id="&user_id
rs.open sql,conn,1,1
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
.style13 {color: #FF0000}
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
                                <td><div align="center"><a href="ynumber_set.asp"><img src="adminimages/numextent.gif" width="134" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="info_search.asp"><img src="adminimages/stuquerry.gif" width="134" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="info_add.asp"><img src="adminimages/infoadd.gif" width="134" height="24" border="0"></a></div></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="86" background="adminimages/titlebk2.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="45"><div align="center"><img src="adminimages/stuinfo.gif" width="523" height="45"></div></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr>
                        <td background="../user/userimages/titlebk.gif"><div align="center">
                            <table width="90%"  border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td colspan="6" valign="top"><div align="center">
                                    <form name="form" method="post" action="info_set1.asp?NoncePage=<%=NoncePage%>">
                                      <table width="100%" border="0" cellpadding="0" cellspacing="0" class="thin">
                                        <tr>
                                          <td width="32%" height="24"><div align="right" class="style10">用户名：&nbsp;&nbsp;</div></td>
                                          <td width="68%" class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_account" value="<%=rs("user_account")%>" size="24" class="style3">
                                              <span class="style13">
 *</span> </td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">密码：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_pwd" size="24" class="style3" value="<%=rs("user_pwd")%>">
                                              <span class="style13">*</span> </td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">姓名：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_name" value="<%=rs("user_name")%>" size="24" class="style3">
                                              <span class="style13">*</span> </td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">学号：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_number" value="<%=rs("user_number")%>" size="24" class="style3">
                                              <span class="style13">*</span> </td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">生日：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                            <input type="text" name="user_Csrq" value="<%=rs("user_Csrq")%>" size="24" class="style3">
                                            （形式：1982-08-09） </td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">专业：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_major" value="<%=rs("user_major")%>" size="24" class="style3"></td>
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
elseif rs("user_grade") = "待添加"  then
%>
                                                <option value="研一" selected>研一</option>
                                                <option value="研二">研二</option>
                                                <option value="研三">研三</option>
                                                <%
												
end if
%>
                                              </select>
                                              <span class="style13">*</span></td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">身份证号码：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_Sfzh" value="<%=rs("user_Sfzh")%>" size="24" class="style3"></td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">考试单位：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_Ksdw" value="<%=rs("user_Ksdw")%>" size="24" class="style3"></td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">考试编号：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_Ksbh" value="<%=rs("user_Ksbh")%>" size="24" class="style3"></td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">报名号：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_Bmh" value="<%=rs("user_Bmh")%>" size="24" class="style3"></td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">考试方式：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_Ksfs" value="<%=rs("user_Ksfs")%>" size="24" class="style3"></td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">毕业单位：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_Bydw" value="<%=rs("user_Bydw")%>" size="24" class="style3"></td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">定向委培单位：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_Dxwpdw" value="<%=rs("user_Dxwpdw")%>" size="24" class="style3"></td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">民族：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_Mz" value="<%=rs("user_Mz")%>" size="24" class="style3"></td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">婚否：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_Hf" value="<%=rs("user_Hf")%>" size="24" class="style3"></td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">国家/地区：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_Gj" value="<%=rs("user_Gj")%>" size="24" class="style3"></td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">E-mail：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_mail" value="<%=rs("user_mail")%>" size="24" class="style3">
                                              <span class="style13">*</span> </td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">BBS帐号：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input name="user_bbs" type="text" class="style3" id="user_bbs" value="<%=rs("user_bbs")%>" size="24">
                                              <span class="style13">*</span> </td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">联系方式：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_roomphone" value="<%=rs("user_roomphone")%>" size="24" class="style3">
                                              <span class="style13">*</span> （形式：025-83594521） </td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">导师：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_tutor" value="<%=rs("user_tutor")%>" size="24" class="style3"></td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">家庭电话：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_homephone" value="<%=rs("user_homephone")%>" size="24" class="style3">
                                （形式：025-83594521） </td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">家庭地址：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_address" value="<%=rs("user_address")%>" size="24" class="style3">                                          </td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">邮编：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input type="text" name="user_code" value="<%=rs("user_code")%>" size="24" class="style3">                                          </td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">性别：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <input name="user_sex" type="radio" value="男" checked>
                                男
                                <input name="user_sex" type="radio" value="女">
                                女 <span class="style13">*</span> </td>
                                        </tr>
                                        <tr>
                                          <td height="24"><div align="right" class="style10">备注：&nbsp;&nbsp;</div></td>
                                          <td class="style10">&nbsp;&nbsp;&nbsp;
                                              <textarea name="user_info" cols="45" rows="10" class="style3"><%=rs("user_info")%></textarea>                                          </td>
                                        </tr>
                                        <tr>
                                          <td height="35" colspan="2"><div align="center"><img src="../user/userimages/editSub.gif" width="70" height="25" style="cursor:hand;" onMousedown="submit()"> </div></td>
                                        </tr>
                                      </table>
                                    </form>
                                </div></td>
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
