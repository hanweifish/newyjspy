<!--#include file="conn.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select top 8 * from notice order by notice_time desc"
rs.open sql,conn,1,1
%>
<%
set rsb=server.createobject("adodb.recordset")
sql="select top 4 * from policy order by policy_time desc"
rsb.open sql,conn,1,1
%>
<script language="javascript">
window.status="欢迎访问南京大学物理系研究生管理信息系统！"
</script>
<script language="javascript">
	function checkuser(form)
	{
		if (document.form.admin_account.value=="")
		{
			alert("请输入用户名！！");
		}
		else if (document.form.admin_pwd.value=="")
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
<!--#include file="top.asp"-->
<table width="800"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="187" height="10"></td>
    <td rowspan="7" valign="top"><div align="right">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="22"><div align="right"></div></td>
        </tr>
        <tr>
          <td valign="top"><div align="right">
            <table width="603"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="406" height="55" background="indeximages/notice.gif">&nbsp;</td>
                <td width="6" background="indeximages/midLinkTop.gif">&nbsp;</td>
                <td background="indeximages/bulletin.gif">&nbsp;</td>
              </tr>
              <tr>
                <td height="25" background="indeximages/noticebktop.gif">&nbsp;</td>
                <td background="indeximages/midLinkTop.gif">&nbsp;</td>
                <td background="indeximages/bulletinbktop.gif">&nbsp;</td>
              </tr>
              <tr>
                <td><table width="100%" height="120" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td background="indeximages/noticeBk.gif"><div align="center">
                          <iframe src="notice.asp" name="notice" width="330" marginwidth="0" height="200" marginheight="0" align="middle" scrolling="yes" frameborder="0" allowtransparency="true"></iframe>
                      </div></td>
                    </tr>
                    <tr>
                      <td height="15" background="indeximages/noticeBk.gif"><div align="center"></div></td>
                    </tr>
                </table></td>
                <td>&nbsp;</td>
                <td valign="middle" background="indeximages/bulletinBk.gif">
                    <div align="center">
                      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="10">&nbsp;</td>
                          <td><div align="center">
<!--#include file="bulletin.asp"-->
						  </div></td>
                        </tr>
                      </table>
					</div></td>
              </tr>
              <tr>
                <td height="60" valign="top" background="indeximages/noticeBt.gif"><div align="center"><a href="noticelist.asp" class="style3" target="_blank">&lt;&lt; MORE &gt;&gt; </a></div></td>
                <td background="indeximages/midLinkBt.gif">&nbsp;</td>
                <td valign="top" background="indeximages/bulletinBt.gif"><div align="center"><a href="policylist.asp" class="style2" target="_blank">&lt;&lt; MORE &gt;&gt; </a></div></td>
              </tr>
            </table>
          </div></td>
        </tr>
        <tr>
          <td height="15" valign="top"><div align="right"></div></td>
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
    <td height="47" background="teacher/adminimages/adminlogin.gif">&nbsp;</td>
  </tr>
  <tr>
    <td><div align="center">
      <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="132"><div align="center">
              <form action="teacher/admin_check.asp" method="post" name="form" id="form">
                <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" background="indeximages/loginbk.gif">
                  <tr>
                    <td colspan="2"><div align="center">教务员:
                            <input name="admin_account" type="text" class="style3" id="name" size="12">
                      </div>
                        <div align="left"> </div></td>
                  </tr>
                  <tr>
                    <td height="38" colspan="2"><div align="center">密 &nbsp;码:
                            <input name="admin_pwd" type="password" class="style3" id="pwd" size="12">
                      </div>
                        <div align="left"> </div></td>
                  </tr>
                  <tr>
                    <td colspan="2"><div align="center"><img src="indeximages/login.gif" width="49" height="23" border="0" align="absmiddle" style='cursor:hand' onMouseDown="checkuser(form)">&nbsp;</div></td>
                  </tr>
                  <tr>
                    <td width="50%" height="10"><div align="center"></div></td>
                    <td width="50%">&nbsp;</td>
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
              <iframe src="denote2.asp" name="denote" width="140" marginwidth="0" height="110" marginheight="0" align="middle" scrolling="yes" frameborder="0" allowtransparency="true" ></iframe>
          </div></td>
        </tr>
      </table>
    </div></td>
  </tr>
  <tr>
    <td height="77" background="indeximages/links.gif">&nbsp;</td>
  </tr>
  <tr>
    <td background="indeximages/loginbk.gif"><div align="center">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
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
      </table>
    </div></td>
  </tr>
  <tr>
    <td background="indeximages/loginbk.gif">&nbsp;</td>
  </tr>
  <tr>
    <td height="34" background="indeximages/loginbottom.gif">&nbsp;</td>
  </tr>
</table>
<!--#include file="bottom.asp"-->
</html>
