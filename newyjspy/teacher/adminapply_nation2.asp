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
dim admin_account
admin_account=session("admin_account")
set rs=server.createobject("adodb.recordset")
sql="select * from apply_nation order by pclb,user_tutor"
rs.open sql,conn,1,1
%>

<script language="javascript">
	function checkform(form3)
	{
		if (document.form3.studentno.value=="")
		{
			alert("请输入学生学号！");
		}
		else
		{
			form3.submit();
		}
		return false;
	}
</script>
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
<%
	if session("popup") = "" then
	set rsnew = server.createobject("adodb.recordset")
	newsql = "select * from guestbook where admin_read = false "
	rsnew.open newsql,conn,1,1
	if not(rsnew.eof or rsnew.bof) then
	Response.write "<script> alert('您有新留言！')</script> "
	end if
	session("popup")="1"
	end if
%>
<table width="800"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="187" height="3"></td>
    <td rowspan="2" valign="top"><div align="right">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="22"><div align="right"></div></td>
        </tr>
        <tr>
          <td valign="top"><div align="right">
            <table width="603"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="86" background="adminimages/titlebk2.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="45"><div align="center"><img src="adminimages/stuinfo.gif" width="523" height="45"></div></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td background="../user/userimages/titlebk.gif"><div align="center">
                    <table width="100%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td colspan="6" valign="middle"><br>
                        <div align="center" class="style10"><br>
                          <%       if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=10
			NumPage=rs.Pagecount
			if request("page")=empty then 
			NoncePage=1
		else
		if Cint(request("page"))<1 then
			NoncePage=1
		else
			NoncePage=Trim(request("page"))
		end if
		if Cint(Trim(request("page")))>Cint(NumPage) then NoncePage=NumPage
	end if
else
	NumRecord=0
	NumPage=0
	NoncePage=0
	end if
%>
                          <table width="90%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                            <tr>
                              <td width="38" height="24"><div align="center" class="style10">姓名</div></td>
                                  <td width="54" height="24"><div align="center" class="style10">学号</div></td>
                                  <td width="24" height="24" align="center"><div align="center" class="style10">专业</div></td>
                                  <td width="38" height="24"><div align="center" class="style10">身份</div></td>
                                  <td width="44"><div align="center" class="style10">出生年月</div></td>
                                  <td width="26"><div align="center" class="style10">导师</div></td>
                                  <td width="40"><div align="center" class="style10">外语成绩</div></td>
                                  <td width="30"><div align="center" class="style10">拟申请学校专业</div></td>
                                  <td width="35"><div align="center" class="style10">派出类别</div></td>
                                  <td width="35"><div align="center" class="style10">计划派出日期</div></td>
                                  <td width="35"><div align="center" class="style10">联系电话</div></td>
                                  <td width="43"><div align="center" class="style10">E_mail</div></td>
                                  <td width="32" height="24"><div align="center" class="style10">修改</div></td>
                                  <td width="39" height="24"><div align="center" class="style10">删除</div></td>
                                </tr>
                            <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*10,1
	for i=1 to rs.pagesize
%>
                            <tr>
                              <td width="38" height="24"><div align="center" class="style10"><a href=name_search.asp?user_name=<%=rs("user_name")%> class="style3"><%=rs("user_name")%></a></div></td>
                                  <td width="54" height="24"><div align="center" class="style10"><%=rs("user_number")%></div></td>
                                  <td height="24"><div align="center" class="style10"><%=rs("user_major")%></div></td>
                                  <td width="38" height="24"><div align="center" class="style10"><%=rs("user_Pyxz")%></div></td>
                                  <td width="44"><div align="center" class="style10"><%=rs("user_Csrq")%></div></td>
                                  <td width="26"><div align="center" class="style10"><%=rs("user_Tutor")%></div></td>
                                  <td width="40"><div align="center" class="style10"><%=rs("sorts")%>,<%=rs("score")%></div></td>
                                  <td width="30"><div align="center" class="style10"><%=rs("school")%>,<%=rs("major")%></div></td>
                                  <td width="35"><div align="center" class="style10"><%=rs("Pclb")%></div></td>
                                  <td width="35"><div align="center" class="style10"><%=rs("applydate")%></div></td>
                                  <td width="35"><div align="center" class="style10"><%=rs("phone")%></div></td>
                                  <td width="43"><div align="center" class="style10"><%=rs("email")%></div></td>
                                  <td width="32" height="24"><div align="center" class="style3"><font color="#ff6633">修改</font></div></td>
                                  <td width="39" height="24"><div align="center" class="style3"><a href=adminapply_nationdel.asp?NoncePage=<%=NoncePage%>&user_number=<%=rs("user_number")%>><font color="#ff6633">删除</font></a></div></td>
                                </tr>
                            <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>没有找到任何记录!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
                            <tr>
                              <td height="40" colspan="14"><div align="center">
                                <input type="hidden" name="page" value="<%=NoncePage%>">
                                <form name="form1" method="post" action="admin_index.asp">
                                  <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                      <td valign="middle"><div align="right">
                                        <div align="right"> <span class="style2">
                                          <%
if NoncePage>1 then
	response.write "|<a href=admin_index.asp?page=1>首 页</a>| |<a href=admin_index.asp?page="&NoncePage-1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_index.asp?page="&NoncePage+1&">下一页</a>| |<a href=admin_index.asp?page="&NumPage&">尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
  &nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录  &nbsp;&nbsp;转到
                                          <input name="page" type="text" class="style3" id="page" size="2">
                                          页</span>&nbsp;&nbsp; </div>
                                                  </div></td>
                                              </tr>
                                    </table>
                                          </form>
                                  </div></td>
                                </tr>
                            </table>
                        </div></td></tr>
                    </table>
                </div></td>
              </tr>
              <tr>
                <td height="40" background="adminimages/titlebk3.gif">&nbsp;</td>
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
