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
dim admin_account
admin_account=session("admin_account")
%>

<%
dim user_name
user_number=trim(request("user_number"))
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_number='"&user_number&"'"
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
.style13 {color: #FF0000}
-->
</style>
<style type="text/css">
<!--
.style12 {color: #006699; font-size: 13px;}
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
                                    <table width="100%"  border="0" cellpadding="0" cellspacing="0">
                                      <%
if not rs.eof then
%>
                                      <%
do while not rs.eof
%>
                                      <tr>
                                        <td><div align="center">
                                            <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                              <tr>
                                                <td width="35%" height="24"><div align="right" class="style10">用户名：&nbsp;&nbsp;</div></td>
                                                <td width="65%" class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_account")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">密码：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_pwd")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">姓名：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_name")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">学号：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_number")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">身份证号码：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_Sfzh")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">考试单位：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_Ksdw")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">考试编号：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_Ksbh")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">报名号：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_Bmh")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">考试方式：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_Ksfs")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">培养性质：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_Pyxz")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">毕业单位：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_Bydw")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">定向委培单位：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_Dxwpdw")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">民族：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_Mz")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">婚否：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_Hf")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">国家/地区：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_Gj")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">专业：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_major")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">院系：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_yx")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right"><span class="style3">毕业信息</span><span class="style10">：&nbsp;&nbsp;</span></div></td>
                                                <td class="style10"><div align="left">&nbsp;&nbsp;&nbsp; 
												<%
													set rsBiyeInfo = Server.CreateObject("Adodb.Recordset")
													sql_rsBiyeInfo = "Select * from biyeInfo where user_number = '"&rs("user_number")&"'"
													rsBiyeInfo.Open sql_rsBiyeInfo,conn,1,1
													
													if rsBiyeInfo.RecordCount = 0 then
												%>
												<a href="AddbiyeInfo.asp?page=<%=NoncePage%>&user_number=<%=rs("user_number")%>" target="_self" class="style3">添加</a>
												<%
													elseif rsBiyeInfo("bylw") = "暂无" then
												%>
												<a href="AddbiyeInfo.asp?page=<%=NoncePage%>&user_number=<%=rs("user_number")%>" target="_self" class="style3">添加</a>
												<%
													else
												%>
												<a href="ViewbiyeInfo.asp?page=<%=NoncePage%>&user_number=<%=rs("user_number")%>" target="_self" class="style3">查看</a>
												<%
													end if
													rsBiyeInfo.Close
												%>
												
												</div></td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">年级：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_grade")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">E-mail：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_mail")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">BBS帐号：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_bbs")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">联系方式：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_roomphone")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">导师：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_tutor")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">家庭电话：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_homephone")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">家庭地址：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_address")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">邮编：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_code")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">性别：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_sex")%></td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">生日：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_Csrq")%> </td>
                                              </tr>
                                              <tr>
                                                <td height="24"><div align="right" class="style10">备注：&nbsp;&nbsp;</div></td>
                                                <td class="style10">&nbsp;&nbsp;&nbsp; <%=HTMLEncode(rs("user_info"))%> </td>
                                              </tr>
                                              <tr>
                                                <td height="35" colspan="2"><div align="center"><strong><a href=info_set.asp?user_ID=<%=rs("user_ID")%> class="style6"><img src="../user/userimages/edit.gif" border="0" class="style2"></a></strong></div></td>
                                              </tr>
                                            </table>
                                        </div></td>
                                      </tr>
                                      <%
rs.movenext
loop
else
%>
                                      <tr>
                                        <td height="25"><div align="center" class="style13 style15"><span class="style5">暂时没有相关信息的记录！</span></div></td>
                                      </tr>
                                      <%
end if
%>
                                    </table>
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
