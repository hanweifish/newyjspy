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
                                <td><div align="center"></div>
                                    <div align="center"><a href="tutor_add.asp"><img src="adminimages/tutorAdd.gif" width="240" height="24" border="0"></a></div>
                                    <div align="center"></div></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="86" background="adminimages/titlebk2.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="45"><div align="center"><img src="../indeximages/tutor.gif" width="523" height="45"></div></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr>
                        <td background="../user/userimages/titlebk.gif"><div align="center">
                            <table width="90%"  border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <th scope="col"><div align="center">
                                    <form name="form1" method="post" action="../admin/tutor_add1.asp">
                                      <table width="90%" height="100%"  border="1" cellpadding="4" cellspacing="0" bordercolor="#000000"class="thin">
                                        <tr>
                                          <td width="60" height="20" bgcolor="#dde6f4"><div align="center" class="style20">导师姓名</div></td>
                                          <td width="158"><div align="center" class="style20">
                                              <div align="left">
                                                <input name="tutor_name" type="text" class="style2" id="tutor_name">
                                              </div>
                                          </div></td>
                                          <td width="67" bgcolor="#dde6f4"><div align="center" class="style20">专业名称</div></td>
                                          <td width="162" class="style20"><div align="left">
                                              <select name="tutor_major" id="tutor_major">
                                                <%
									  if session("tutor_major")="" or session("tutor_major")="理论物理" then
									  %>
                                                <option value="理论物理" selected>理论物理</option>
                                                <option value="粒子物理与原子核物理">粒子物理与原子核物理</option>
                                                <option value="凝聚态物理">凝聚态物理</option>
                                                <option value="光学">光学</option>
                                                <option value="生物物理学">生物物理学</option>
                                                <option value="微电子学与固体电子学">微电子学与固体电子学</option>
                                                <%
									  elseif session("tutor_major")="粒子物理与原子核物理" then
									  %>
                                                <option value="理论物理">理论物理</option>
                                                <option value="粒子物理与原子核物理" selected>粒子物理与原子核物理</option>
                                                <option value="凝聚态物理">凝聚态物理</option>
                                                <option value="光学">光学</option>
                                                <option value="生物物理学">生物物理学</option>
                                                <option value="微电子学与固体电子学">微电子学与固体电子学</option>
                                                <%
									  elseif session("tutor_major")="凝聚态物理" then
									  %>
                                                <option value="理论物理">理论物理</option>
                                                <option value="粒子物理与原子核物理">粒子物理与原子核物理</option>
                                                <option value="凝聚态物理" selected>凝聚态物理</option>
                                                <option value="光学">光学</option>
                                                <option value="生物物理学">生物物理学</option>
                                                <option value="微电子学与固体电子学">微电子学与固体电子学</option>
                                                <%
									  elseif session("tutor_major")="光学" then
									  %>
                                                <option value="理论物理">理论物理</option>
                                                <option value="粒子物理与原子核物理">粒子物理与原子核物理</option>
                                                <option value="凝聚态物理">凝聚态物理</option>
                                                <option value="光学" selected>光学</option>
                                                <option value="生物物理学">生物物理学</option>
                                                <option value="微电子学与固体电子学">微电子学与固体电子学</option>
                                                <%
									  elseif session("tutor_major")="生物物理学" then
									  %>
                                                <option value="理论物理">理论物理</option>
                                                <option value="粒子物理与原子核物理">粒子物理与原子核物理</option>
                                                <option value="凝聚态物理">凝聚态物理</option>
                                                <option value="光学">光学</option>
                                                <option value="生物物理学" selected>生物物理学</option>
                                                <option value="微电子学与固体电子学">微电子学与固体电子学</option>
                                                <%
									  elseif session("tutor_major")="微电子学与固体电子学" then
									  %>
                                                <option value="理论物理">理论物理</option>
                                                <option value="粒子物理与原子核物理">粒子物理与原子核物理</option>
                                                <option value="凝聚态物理">凝聚态物理</option>
                                                <option value="光学">光学</option>
                                                <option value="生物物理学">生物物理学</option>
                                                <option value="微电子学与固体电子学" selected>微电子学与固体电子学</option>
                                                <%
										end if
										%>
                                              </select>
                                          </div></td>
                                        </tr>
                                        <tr>
                                          <td width="60" height="20" bgcolor="#dde6f4"><div align="center" class="style20">职&nbsp;&nbsp;&nbsp; 称 </div></td>
                                          <td><div align="center" class="style20">
                                              <div align="left">
                                                <select name="tutor_post" id="tutor_post">
                                                  <option value="教授" selected>教授</option>
                                                  <option value="副教授">副教授</option>
                                                </select>
                                              </div>
                                          </div></td>
                                          <td bgcolor="#dde6f4"><div align="center" class="style20">博导/硕导</div></td>
                                          <td width="162"><div align="left"><span class="style20">
                                              <select name="tutor_mord" id="tutor_mord">
                                                <option value="博导" selected>博导</option>
                                                <option value="硕导">硕导</option>
                                              </select>
                                          </span></div></td>
                                        </tr>
                                        <tr>
                                          <td width="60" height="20" bgcolor="#dde6f4"><div align="center" class="style20">是否院士</div></td>
                                          <td><div align="center" class="style20">
                                              <div align="left">
                                                <select name="tutor_acad" id="tutor_acad">
                                                  <option value="是">是</option>
                                                  <option value="否" selected>否</option>
                                                </select>
                                              </div>
                                          </div></td>
                                          <td bgcolor="#dde6f4"><div align="center" class="style20">是否兼职博导</div></td>
                                          <td width="162"><div align="left">
                                              <select name="tutor_ptj" id="tutor_ptj">
                                                <option value="是">是</option>
                                                <option value="否" selected>否</option>
                                              </select>
                                          </div></td>
                                        </tr>
                                        <tr bgcolor="#dde6f4">
                                          <td height="25" colspan="4"><div align="left"><strong><span class="style20">&nbsp;&nbsp;学科专长及研究方向：</span></strong></div></td>
                                        </tr>
                                        <tr>
                                          <td colspan="4"><div align="left"><span class="style20"> &nbsp;&nbsp;
                                                    <textarea name="tutor_dir" cols="50" rows="5" wrap="physical"></textarea>
                                          </span></div></td>
                                        </tr>
                                        <tr bgcolor="#dde6f4">
                                          <td height="25" colspan="4"><div align="left"><strong><span class="style20">&nbsp;&nbsp;科研项目：</span></strong></div></td>
                                        </tr>
                                        <tr>
                                          <td colspan="4"><div align="left"><span class="style20"><br>
&nbsp;&nbsp;
                                  <textarea name="tutor_proj" cols="50" rows="5"></textarea>
                                  <br>
                                          </span></div></td>
                                        </tr>
                                        <tr>
                                          <td height="24" colspan="4"><div align="center"><img src="../user/userimages/add.gif" width="51" height="23" align="absmiddle" style="cursor:hand " onMouseDown="javascript:submit();"> </div></td>
                                        </tr>
                                      </table>
                                    </form>
                                </div></th>
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
