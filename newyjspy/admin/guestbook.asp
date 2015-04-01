<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("admin_account")="" or session("user_group")="subadmin" then
Response.write"对不起，您还没有登陆，无此权限！"
Response.end
end if
%>

<%
dim admin_account
admin_account=session("admin_account")
%>

<%
set rs=server.createobject("adodb.recordset")
sql="select user_info.user_name,user_info.user_mail,user_info.user_ID,guestbook.admin_time,guestbook.guestbook_ID,guestbook.admin_read,guestbook.user_time,guestbook.title from guestbook inner join user_info on guestbook.user_ID=user_info.user_ID order by guestbook.user_time desc"
rs.open sql,conn,1,1
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
.style12 {color: #006699;
	font-size: 12px;
}
-->
</style>
<style type="text/css">
<!--
.style14 {color: #FF6633;
	font-size: 12px;
}
.style15 {font-size: 11px}
-->
</style>
<style type="text/css">
<!--
.style16 {color: #FF6600}
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
                                <td><div align="center"><img src="adminimages/messagemag.gif" width="523" height="45"></div></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="86" background="adminimages/titlebk2.gif"><div align="center" class="style2">选 择 发 送 对 象</div></td>
                      </tr>
                      <tr>
                        <td background="../user/userimages/titlebk.gif"><div align="center">
                            <table width="100%"  border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td><div align="center">
                                    <%if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=10
			NumPage=rs.Pagecount
			if request("page")=empty then 
			NoncePage=1
		else
		if Cint(request("page"))<1 then
			NoncePage=1
		else
			NoncePage=request("page")
		end if
		if Cint(Trim(request("page")))>Cint(NumPage) then NoncePage=NumPage
	end if
else
	NumRecord=0
	NumPage=0
	NoncePage=0
	end if
%>
                                    <table width="100%"  cellspacing="0" cellpadding="0">
                                      <tr>
                                        <td><div align="center">
                                            <table width="90%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                              <tr align="center">
                                                <td width="13%" height="24" align="center" valign="middle"><div align="center" class="style10"> 发送时间 </div></td>
                                                <td width="6%">&nbsp;</td>
                                                <td width="41%" height="24"><div align="center" class="style10">标 &nbsp;&nbsp;&nbsp;题 </div></td>
                                                <td width="13%"><span class="style10">浏览时间</span></td>
                                                <td width="11%" class="style10">留言人</td>
                                                <td width="8%" class="style10">删除</td>
                                                <td width="8%" class="style10">回复</td>
                                              </tr>
<%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*10,1
	for i=1 to rs.pagesize
%>
                                              <%if rs("admin_read")= true then %>
                                              <tr>
                                                <td height="25" valign="middle" align="center"><div align="left" class="style2">
                                                    <div align="center"><%=rs("user_time")%></div>
                                                </div></td>
                                                <td>&nbsp;</td>
                                                <td height="25"><div align="left"></div>
                                                    <div align="left"class="style2">&nbsp;&nbsp;<a href=guestbook_detail.asp?guestbook_ID=<%=rs("guestbook_ID")%>><%=rs("title")%></a></div></td>
                                                <td><div align="center"class="style2"><%=rs("admin_time")%></div></td>
                                                <td><div align="center"class="style2"><%=rs("user_name")%></div></td>
                                                <td><div align="center"><a href=guestbook_del.asp?guestbook_ID=<%=rs("guestbook_ID")%> class="style6" >删除</a></div></td>
                                                <td><div align="center" class="style6"><a href=guestbook_reply1.asp?user_ID=<%=rs("user_ID")%>&title=<%=rs("title")%> class="style6" >回复</a></div></td>
                                              </tr>
                                              <%else%>
                                              <tr>
                                                <td height="25" valign="middle" align="center"><div align="left" class="style2">
                                                    <div align="center"><%=rs("user_time")%></div>
                                                </div></td>
                                                <td><span class="style8 style13"><img src="adminimages/new.GIF" width="28" height="11" align="absmiddle"></span></td>
                                                <td height="25"><div align="left"></div>
                                                    <div align="left"class="style8 style13">&nbsp;&nbsp;&nbsp;<a href=guestbook_detail.asp?guestbook_ID=<%=rs("guestbook_ID")%> class="style2"><%=rs("title")%></a></div></td>
                                                <td><div align="center"class="style2"><%=rs("admin_time")%></div></td>
                                                <td class="style3"><div align="center"class="style2"><%=rs("user_name")%></div></td>
                                                <td><div align="center"><a href=guestbook_del.asp?guestbook_ID=<%=rs("guestbook_ID")%> class="style6">删除</a></div></td>
                                                <td><div align="center" class="style6"><a href=guestbook_reply1.asp?user_ID=<%=rs("user_ID")%>&title=<%=rs("title")%> class="style6" >回复</a></div></td>
                                              </tr>
                                              <%end if%>
                                              <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='25'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>暂时还没有信件!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
                                              <tr>
                                                <td height="24" colspan="13" align="center" valign="middle"><div align="right"> <span class="style2">
                                                    <input type="hidden" name="page" value="<%=NoncePage%>">
                                                    <%
if NoncePage>1 then
	response.write "|<a href=guestbook.asp?page=1>首 页</a>| |<a href=guestbook.asp?page="&NoncePage-1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=guestbook.asp?page="&NoncePage+1&">下一页</a>| |<a href=guestbook.asp?page="&NumPage&">尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
                                                </span> <span class="style10">&nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录</span>&nbsp; </div></td>
                                              </tr>
                                            </table>
                                        </div></td>
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
<!--#include file="bottom1.asp"-->
</html>
