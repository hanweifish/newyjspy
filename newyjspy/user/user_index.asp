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

<!--#include file="regfirst.asp"--> 
 
<%
set rsn=server.createobject("adodb.recordset")
sql="select top 8 * from notice where notice_authority='private' order by notice_time desc"
rsn.open sql,conn,1,1
%>
<%
set rsb=server.createobject("adodb.recordset")
sql="select top 4 * from policy order by policy_time desc"
rsb.open sql,conn,1,1
%>
<script language="javascript">
window.status="欢迎访问南京大学物理系研究生管理信息系统！"
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>研究生信息管理系统</title>
<link href="../style.css" rel="stylesheet" type="text/css">
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
                <td height="10">&nbsp;</td>
                <td background="../indeximages/midLinkTop.gif">&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td width="406" height="55"  background="../indeximages/notice.gif">&nbsp;</td>
                <td width="6" background="../indeximages/midLinkTop.gif">&nbsp;</td>
                <td background="../indeximages/bulletin.gif">&nbsp;</td>
              </tr>
              <tr>
                <td height="25" background="../indeximages/noticebktop.gif">&nbsp;</td>
                <td background="../indeximages/midLinkTop.gif">&nbsp;</td>
                <td background="../indeximages/bulletinbktop.gif">&nbsp;</td>
              </tr>
              <tr>
                <td><table width="100%" height="90" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td background="../indeximages/noticeBk.gif"><div align="center">
                          <iframe src="notice1.asp" name="notice1" width="330" marginwidth="0" height="200" marginheight="0" align="middle" scrolling="yes" frameborder="0" allowtransparency="true"></iframe>
                      </div></td>
                    </tr>
                </table></td>
                <td>&nbsp;</td>
                <td background="../indeximages/bulletinBk.gif"><div align="center">
                    <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="top"><div align="center">
                            <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td valign="top"><div align="center">
                                    <marquee direction="up" width="161" height="200" loop="-1" scrollamount="1" scrolldelay="1">
                                    <table width="85%" border="0" cellpadding="0" cellspacing="0" bordercolor="#A8BAFF">
                                      <%
if not (rsb.eof and rsb.bof) then
for i=1 to rsb.recordcount
%>
                                      <tr>
                                        <td height="5"></td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                      </tr>
                                      <tr>
                                        <td width="9%" height="28">&nbsp;</td>
                                        <td width="8%"><div align="center"><img src="../indeximages/triangle.gif" width="6" height="8" align="middle"></div></td>
                                        <td width="83%"><font class="style2"><a href=policy_detail.asp?policy_ID=<%=rsb("policy_ID")%> target="_blank" class="style3"><%=rsb("policy_title")%></a></font></td>
                                        <%
rsb.movenext
next
%>
                                      </tr>
                                    </table>
                                    </marquee>
                                </div></td>
                              </tr>
                            </table>
                            <%
else
%>
                            <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#A8BAFF">
                              <tr>
                                <td><div align="center"><font class="style3" >暂时没有新的政策发布！</font></div></td>
                              </tr>
                            </table>
                            <%
	end if
%>
<%
rsb.close
set rsb=nothing
%>
                        </div></td>
                      </tr>
                    </table>
                </div></td>
              </tr>
              <tr>
                <td height="60" background="../indeximages/noticeBt.gif">&nbsp;</td>
                <td background="../indeximages/midLinkBt.gif">&nbsp;</td>
                <td background="../indeximages/bulletinBt.gif">&nbsp;</td>
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
<%
rs.close
set rs=nothing
%>
<!--#include file="bottom1.asp"-->
</html>
