<!--#include file="conn.asp"-->

<%
dim job_ID
job_ID=trim(request("job_ID"))
set rs=server.createobject("adodb.recordset")
sql="select * from job where job_ID="&job_ID
rs.open sql,conn,1,3
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
<%
	rs("job_click")=rs("job_click")+1
	rs.update
%>
<script language="javascript">
window.status="欢迎访问南京大学物理系研究生管理信息系统！"
</script>

<script language="javascript">
	function checkuser(form)
	{
		if (document.form.user_account.value=="")
		{
			alert("请输入用户名！！");
		}
		else if (document.form.user_pwd.value=="")
		{
			alert("请输入密码！!");
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>研究生信息管理系统</title>
<link href="style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style10 {font-size: 12px;
	color: #004080;
}
-->
</style>
<style type="text/css">
<!--
.style12 {color: #FF0000}
-->
</style>
<style type="text/css">
<!--
.style9 {font-size: 12px}
-->
</style>
<!--#include file="top.asp"-->
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
                        <td height="54" background="user/userimages/titlebk1.gif"><div align="center"><img src="indeximages/job.gif" width="523" height="45"></div></td>
                      </tr>
                      <tr>
                        <td height="53" background="user/userimages/titlebk2.gif">&nbsp;</td>
                      </tr>
                      <tr>
                        <td background="user/userimages/titlebk.gif"><div align="center">
                            <table width="90%"  border="0" cellpadding="0" cellspacing="0" class="thin">
                              <tr>
                                <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                      <td height="25" bgcolor="#CFDBEF"><div align="center" class="style3"><span class="style15 style21"><%=rs("job_title")%></span></div></td>
                                    </tr>
                                    <tr>
                                      <td><div align="center">
                                          <table width="90%"  border="0" cellspacing="0" cellpadding="0">
                                            <tr>
                                              <td width="39%" height="20" class="style10"><span class="style2">发布人：</span><%=rs("job_author")%></td>
                                              <td width="17%" class="style9"><div align="right"></div></td>
                                              <td width="44%" class="style9"><div align="right" class="style10"><span class="style2">发布时间：</span><span class="style3"><%=rs("job_time")%></span>&nbsp;&nbsp;</div></td>
                                            </tr>
                                            <tr>
                                              <td height="20" colspan="3"><div align="right"><span class="style10">点击次数：<%=rs("job_click")%>&nbsp;&nbsp;</span></div></td>
                                            </tr>
                                            <tr>
                                              <td colspan="3"><br>
                                                  <span class="style2"><%=HTMLEncode(rs("job_content"))%></span><br>
&nbsp;</td>
                                            </tr>
                                            <tr valign="middle">
                                              <td height="20" colspan="3" valign="middle"><p class="style2">备注信息：</p></td>
                                            </tr>
                                            <tr valign="middle">
                                              <td colspan="3" valign="middle" class="style2" ><div align="left"> <%=HTMLEncode(rs("job_info"))%> </div></td>
                                            </tr>
                                            <tr>
                                              <td height="35" colspan="3"><div align="center"><a href="job.asp"><img src="user/userimages/return.gif" width="49" height="23" border="0"></a></div></td>
                                            </tr>
                                          </table>
                                      </div></td>
                                    </tr>
                                </table></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="34" background="user/userimages/titlebk3.gif">&nbsp;</td>
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
            <!--#include file="server.asp" -->
		  </div></td>
        </tr>
      </table>
    </div></td>
  </tr>
  <tr>
    <td valign="top" background="indeximages/loginbk.gif"><div align="center">
      </div>      <div align="center">
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="47" background="indeximages/stulogin.gif">&nbsp;</td>
          </tr>
          <tr>
            <td><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="132"><div align="center">
                    <form action="user/user_check.asp" method="post" name="form" id="form">
                      <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" background="indeximages/loginbk.gif">
                        <tr>
                          <td colspan="2"><div align="center">用户名:
                                  <input name="user_account" type="text" class="style3" id="name" size="12">
                            </div>
                              <div align="left"> </div></td>
                        </tr>
                        <tr>
                          <td height="38" colspan="2"><div align="center">密 &nbsp;码:
                                  <input name="user_pwd" type="password" class="style3" id="pwd" size="12">
                            </div>
                              <div align="left"> </div></td>
                        </tr>
                        <tr>
                          <td width="50%"><div align="right"><img src="indeximages/login.gif" width="49" height="23" border="0" style='cursor:hand' onMouseDown="checkuser(form)">&nbsp;</div></td>
                          <td width="50%"><div align="left">&nbsp;<a href="javascript:void(null)"><img src="indeximages/register.gif" width="49" height="23" border="0"></a></div></td>
                        </tr>
                        <tr>
                          <td height="10"><div align="center"></div></td>
                          <td>&nbsp;</td>
                        </tr>
                      </table>
                    </form>
                </div></td>
              </tr>
              <tr>
                <td height="6" background="indeximages/loginbk.gif"><div align="center"><img src="indeximages/loginbar.gif" width="129" height="2"></div></td>
              </tr>
              <tr>
                <td height="120" valign="center" background="indeximages/loginbk.gif"><div align="center">
                    <iframe src="denote.asp" name="denote" width="150" marginwidth="0" height="120" marginheight="0" align="middle" scrolling="yes" frameborder="0" allowtransparency="true" ></iframe>
                </div></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td height="77" background="indeximages/links.gif">&nbsp;</td>
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
                <td height="35"><div align="center"><a href="links.asp" class="style3">&gt;&gt;&gt; More</a></div></td>
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
    <td height="34" background="indeximages/loginbottom.gif">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>
<!--#include file="bottom.asp"-->
</html>
