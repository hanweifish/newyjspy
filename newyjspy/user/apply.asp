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

<!--#include file="regfirst.asp"--> 

<%
user_nmuber=session("user_nmuber")
session("user_nmuber")=user_nmuber
set rsn=server.createobject("adodb.recordset")
sql="select * from notice order by notice_time desc"
rsn.open sql,conn,1,1
%>

<html>
<head>
<script language="javascript">
<!--

window.status="欢迎访问南京大学物理系研究生管理信息系统！"
//-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>研究生院信息管理系统</title>
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
<style type="text/css">
<!--
.STYLE15 {font-size: 14pt}
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
                            <td height="54" background="userimages/titlebk1.gif"><div align="center"><img src="userimages/apply.gif" ></div></td>
                          </tr>
                          <tr>
                            <td height="53" background="userimages/titlebk2.gif">&nbsp;</td>
                          </tr>
                          <tr>
                            <td background="userimages/titlebk.gif"><div align="center">
                              
                              <table width="315" height="148" border="1" cellpadding="1" cellspacing="1">
                                <tr>
                                  <td width="100"><div align="center">公派出国申报</div></td>
                                  <td width="100"><div align="center">延期毕业申请</div></td>
                                  <td><div align="center"><a href="apply_zzsbldsq.asp">终止硕博连读申请</a></div></td>
                                </tr>
                                <tr>
                                  <td><div align="center"><a href="confirmbyinfo.asp">毕业信息确认</a></div></td>
                                  <td><div align="center"><a href="apply_tqby.asp">提前毕业申请</a></div></td>
                                  <td><div align="center"><a href="apply_zzy.asp">转专业申请</a></div></td>
                                </tr>
                                <tr>
                                  <td><div align="center"><a href="apply_date.asp">出国日期填写</a>&nbsp;</div></td>
                                  <td><div align="center"><a href="apply_tx.asp">退学申请</a></div></td>
                                  <td><div align="center"><a href="apply_zds.asp">转导师申请</a></div></td>
                                </tr>
                                <tr>
                                  <td><div align="center">&nbsp;</div></td>
                                  <td><div align="center"><a href="apply_xx.asp">休学申请</a></div></td>
                                  <td><div align="center"><a href="apply_qxxj.asp">取消学籍申请</a></div></td>
                                </tr>
                                <tr>
                                  <td>&nbsp;</td>
                                  <td><div align="center"><a href="apply_fx.asp">复学申请</a></div></td>
                                  <td><div align="center">&nbsp;</div></td>
                                </tr>
                              </table>
                              <p>&nbsp;</p>
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
            <td height="15" valign="top"><div align="right"> </div></td>
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
    <td valign="top"  background="../indeximages/loginbk.gif"><div align="center"> </div>
        <div align="center">
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
              <td><table width="100%"  border="0" cellspacing="0" cellpadding="0"  background="../indeximages/loginbk.gif">
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
                                    <option value="javascript:void(null);" selected>----校外链接----</option>
                                    <option value="http://www.njbys.com/">南京毕业生就业网</option>
                                    <option value="http://www.jsbys.com.cn/index.aspx">江苏毕业生就业网</option>
                                    <option value="http://www.firstjob.com.cn/">上海毕业生就业网</option>
                                  </select>
                                </form>
                            </div></td>
                          </tr>
						  <tr>
                  <td height="50"><div align="center">
                      <form name="links">
                        <select name="links" class="style2" onChange="window.open(this.value)">
                          <option value="javascript:void(null);" selected>----校内链接----</option>
                          <option value="http://www.nju.edu.cn/">南京大学</option>
                          <option value="http://lily.nju.edu.cn/">南大小百合</option>
                          <option value="http://grawww.nju.edu.cn/">研究生院</option>
                          <option value="http://physics.nju.edu.cn/">物理学系</option>
                          <option value="http://job.nju.edu.cn/">就业指导中心</option>
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
              <td  background="../indeximages/loginbk.gif">&nbsp;</td>
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
