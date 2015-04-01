<style type="text/css">
<!--
.style10 {font-size: 12px;
	color: #004080;
}
.style13 {color: #FF0000}
.style15 {	color: #FF6600;
	font-size: 15px;
}
.style9 {font-size: 12px}
.style9 {color: #FFFF00}
.style17 {color: #FF6633}
.style12 {color: #006699; font-size: 13px;}
.style16 {color: #FF6633; font-weight: bold; }
.style18 {color: #006699;
	font-size: 12px;
}
.style14 {color: #FF6633;
	font-size: 12px;
}
.style19 {font-size: 11px}
-->
</style>
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
                                <td><div align="center"><a href="sel_search.asp"><img src="adminimages/querryresult.gif" width="134" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="course_set.asp"><img src="adminimages/selcourse.gif" width="134" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="course_add.asp"><img src="adminimages/courseadd.gif" width="134" height="24" border="0"></a></div></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="86" background="adminimages/titlebk2.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="45"><div align="center"><img src="adminimages/selset.gif" width="523" height="45"></div></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr>
                        <td valign="top" background="../user/userimages/titlebk.gif"><div align="center">
                            <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td><div align="center">
                                    <table width="90%" cellpadding="0" cellspacing="0" >
                                      <tr>
                                        <td><form name="form1" method="post" action="sel_result.asp">
                                            <div align="center"> <br>
                                                <br>
                                                <table width="100%" border="1" cellpadding="0"  cellspacing="0" bordercolor="#000000" class="thin">
                                                  <tr valign="bottom">
                                                    <td width="34%" height="25"><div align="right" class="style2">按&nbsp; </div></td>
                                                    <td width="66%" height="25"><div align="left" class="style12">
                                                        <div align="left" class="style2">
                                                          <select name="search_class" class="style3">
                                                            <option value="学生姓名" selected>学生姓名</option>
                                                            <option value="学生学号">学生学号</option>
                                                          </select>
                                            查询</div>
                                                    </div></td>
                                                  </tr>
                                                  <tr valign="bottom">
                                                    <td height="25"><div align="right" class="style2">请输入关键字：</div></td>
                                                    <td><div align="left">
                                                        <input name="keywords" type="text" class="style3" size="30">
&nbsp;&nbsp;&nbsp; <img src="../user/userimages/search.gif" width="46" height="23" align="absmiddle" style="cursor:hand; " onMouseDown="checkform1(form1)"> </div></td>
                                                  </tr>
                                                </table>
                                            </div>
                                        </form></td>
                                      </tr>
                                    </table>
                                </div></td>
                              </tr>
                              <tr>
                                <td><div align="center">
                                    <table width="90%" cellpadding="0" cellspacing="0" >
                                      <tr>
                                        <td><form name="form" method="post" action="sel_result.asp">
                                            <div align="center"> <br>
                                                <br>
                                                <table width="100%" border="1" cellpadding="0"  cellspacing="0" bordercolor="#000000" class="thin">
                                                  <tr valign="bottom">
                                                    <td width="34%" height="25"><div align="right" class="style2">按&nbsp; </div></td>
                                                    <td width="66%" height="25"><div align="left" class="style12">
                                                        <div align="left" class="style2">
                                                          <select name="search_class" class="style3">
                                                            <option value="课程" selected>课程</option>
                                                          </select>
                                            查询</div>
                                                    </div></td>
                                                  </tr>
                                                  <tr valign="bottom">
                                                    <td height="25"><div align="right" class="style2">请输入关键字：</div></td>
                                                    <td><div align="left">
                                                        <select name="keywords" class="style3">
                                                          <option value="----选课课程----">----选课课程----</option>
                                                          <%
											  	set rsS = Server.CreateObject("Adodb.recordset")
												sqlS = "select * from course where course_term ='1' order by course_name"
												rsS.open sqlS,conn,1,1
												if not (rsS.bof and rsS.eof) then
												for i=1 to rsS.recordcount
											  %>
                                                          <option value="<%=rsS("course_name")%>"><%=rsS("course_name")%></option>
                                                          <%
											  	rsS.movenext
												if rsS.eof then exit for
												next
												end if
												rsS.close
												set rsS=nothing
											  %>
                                                        </select>
&nbsp;&nbsp;&nbsp; <img src="../user/userimages/search.gif" width="46" height="23" align="absmiddle" style="cursor:hand; " onMouseDown="checkform(form)"> </div></td>
                                                  </tr>
                                                </table>
                                            </div>
                                        </form></td>
                                      </tr>
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
