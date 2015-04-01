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
dim user_number
user_number=session("user_number")
set rs=server.createobject("adodb.recordset")
sql="select * from confirmbiyeinfo where user_number='"&user_number&"'"
rs.open sql,conn,1,1
%>

<%
dim user_account
user_account=session("user_account")
set rsn=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rsn.open sql,conn,1,1
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
	function checkuser()
	{
		var msg = "";
		if(document.form.phone.value.length < 8||document.form.phone.value == "待添加")
			{
				msg = msg + "    请输入联系方式!\n\n";
			}
		
		if(msg !="") 
			{
				alert(msg);
				return false;
			}
		document.form.submit();
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
<style type="text/css">
<!--
.STYLE14 {color: #FF9900}
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
        <td height="47" background="../indeximages/stulogin.gif">&nbsp;</td>
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
                    <td width="60%" class="style3"><%=rsn("user_account")%>：</td>
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
              <td height="6" background="../indeximages/loginbk.gif"><div align="center"><img src="../indeximages/loginbar.gif" width="129" height="2"></div></td>
            </tr>
            <tr>
              <td height="60" valign="center" background="../indeximages/loginbk.gif"><div align="center"><a href="user_logout.asp"><img src="../includeimages/logout.gif" width="60" height="24" border="0"></a></div></td>
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
                    <td height="54" background="userimages/titlebk1.gif"><div align="center"><img src="userimages/stuinfo.gif" width="523" height="45"></div></td>
                  </tr>
                  <tr>
                    <td height="53" background="userimages/titlebk2.gif">&nbsp;</td>
                  </tr>
                  <tr>
                    <td background="userimages/titlebk.gif"><div align="center">
                      <form name="form" method="post" action="apply_yqby1.asp">
                        <div align="center">
<% 
	if rs.eof then

%>
	<table width="90%" height="100%" border="0" cellpadding="0" cellspacing="0" class="thin">
                            <tr>
                              <td ><div align="center" class="style10">对不起,毕业信息里没有您的信息!!毕业信息确认只限即将毕业的学生。</div></td>
							  </tr>
							  <tr>
                              <td height="35" colspan="3"><div align="center"><a href="apply.asp"><img src="userimages/return.gif" width="49" height="23" border="0"></a><br>
                              </div></td>
                            </tr>
							  </table>
<%
else
%>				
                          <table width="90%" height="100%" border="0" cellpadding="0" cellspacing="0" class="thin">
                            <tr>
                              <td width="18%" height="24"><div align="right" class="style10">姓名：&nbsp;&nbsp;</div></td>
                              <td width="34%" class="style10"><%=rs("user_name")%>                                  </td>
                              </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">学号：&nbsp;&nbsp;</div></td>
                              <td class="style10"><%=rs("user_number")%></td>
                              </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">身份证号：&nbsp;&nbsp;</div></td>
                              <td class="style10"><%=rs("user_sfzh")%></td>
                              </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">性别：&nbsp;&nbsp;</div></td>
                              <td class="style10"><%=rs("user_sex")%></td>
                              </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">出生日期：&nbsp;&nbsp;</div></td>
                              <td class="style10"><%=rs("user_csrq")%></td>
                              </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">专业：&nbsp;&nbsp;</div></td>
                              <td class="style10"><%=rs("user_major")%></td>
                              </tr>
							  <tr>
							    <td height="24"><div align="right" class="style10">所在院系：&nbsp;&nbsp;</div></td>
							    <td  class="style10"><%=rs("user_yx")%></td>
							    </tr>
							  <tr>
                              <td height="24"><div align="right" class="style10">培养性质：&nbsp;&nbsp;</div></td>
                              <td class="style10"><%=rs("user_Pyxz")%></td>
                              </tr>
                            
                              <tr>
                                <td height="24"><div align="right" class="style10">入学日期：&nbsp;&nbsp;</div></td>
                                <td colspan="2" class="style10"><%=rs("user_rxrq")%></td>
                              </tr>
                              <tr>
                                <td height="32"><div align="right" class="style10">入学季节：&nbsp;&nbsp;</div></td>
                                <td  class="style10"><%=rs("user_rxjj")%></td>
                              </tr>
                           
                            
                            
                            <tr>
                              <td height="35" colspan="3"><div align="center"><a href="apply.asp"><img src="userimages/return.gif" width="49" height="23" border="0"></a><br>
                                <br>
                                  （<span class="STYLE14">毕业信息如有错误,请与所在院系教务员联系</span>）<br>
                              </div></td>
                            </tr>
                          </table>
<%
end if
%>
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
<%
rs.close
set rs=nothing

rsn.close
set rsn=nothing
%>
</html>
