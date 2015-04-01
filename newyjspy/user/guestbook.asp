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
set rsu=server.createobject("adodb.recordset")
sql="select * from user_info where user_account = '"&session("user_account")&"'"
rsu.open sql,conn,1,1
%>
<%
set rs1=server.createobject("adodb.recordset")
sql="select * from reply where user_ID = "&rsu("user_ID")
rs1.open sql,conn,1,1
%>


<!--#include file="regfirst.asp"--> 

<script language="javascript">
	function checkguestbook()
	{
		if (document.form.title.value=="")
		{
			alert("标题不能为空！！");
		}
		else if (document.form.content.value=="")
		{
			alert("不能发送无内容的留言！!");
		}
		else
		{
			return true;
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
<style type="text/css">
<!--
.style14 {color: #006699;
	font-size: 12px;
}
.style14 {color: #006699;
	font-size: 13px;
	font-weight: bold;
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
                        <td height="54" background="userimages/titlebk1.gif"><div align="center"><img src="userimages/messages.gif" width="523" height="45"></div></td>
                      </tr>
                      <tr>
                        <td height="53" background="userimages/titlebk2.gif">&nbsp;</td>
                      </tr>
                      <tr>
                        <td background="userimages/titlebk.gif"><div align="center">
						  <table width="90%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="25"><div align="center"><a href="guestbook_send.asp" class="style3">发 表 留 言</a>&nbsp;&nbsp;&nbsp;&nbsp;</div></td>
                            </tr>
                            <tr>
                              <td>
							  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                <tr align="center">
                                  <td width="41%" height="24"><div align="center" class="style10">标 &nbsp;&nbsp;&nbsp;题 </div></td>
                                  <td width="11%" class="style10">留言人</td>
                                  <td width="8%" class="style10">删除</td>
                                  </tr>
<%
if rs1.bof or rs1.eof then
	response.write "<tr><td colspan=13 height='25'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>无回复!!!</font></marquee></td></tr>"
else
	for i=1 to rs1.recordcount
%>
                                <tr>
                                  <td height="25"><div align="left"></div>
                                      <div align="left"class="style2">&nbsp;&nbsp;<a href=guestbook_detail.asp?reply_ID=<%=rs1("reply_ID")%>><%=rs1("title")%></a></div></td>
                                  <td><div align="center"class="style2"><%=rs1("author")%></div></td>
                                  <td><div align="center"><a href=guestbook_del.asp?reply_ID=<%=rs1("reply_ID")%> class="style6" >删除</a></div></td>
                                  </tr>
<%
rs1.movenext
if rs1.eof then exit for
	next
end if	
rs1.close
set rs1=nothing
%>
                              </table></td>
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
                        <td width="60%" class="style3"><%=rsu("user_account")%>：</td>
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
rsu.close
set rsu=nothing
%>
<!--#include file="bottom1.asp"-->
</html>
