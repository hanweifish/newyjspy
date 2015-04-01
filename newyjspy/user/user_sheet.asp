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
set rss=server.createobject("adodb.recordset")
sql="select sheet.course,sheet.score,sheet.property,sheet.sheet_info,subject.credit,subject.term from user_info inner join (sheet inner join subject on sheet.course = subject.course) on user_info.user_ID = sheet.user_ID where user_account ='"&session("user_account")&"' order by subject.term,sheet.score desc"
rss.open sql,conn,1,1
%>

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
<style type="text/css">
<!--
.style9 {font-size: 12px}
-->
</style>
<style type="text/css">
<!--
.style17 {color: #FF6633}
.style18 {	font-size: 15px;
	color: #004080;
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
                        <td height="54" background="userimages/titlebk1.gif"><div align="center"><img src="userimages/sheet.gif"></div></td>
                      </tr>
                      <tr>
                        <td height="53" background="userimages/titlebk2.gif">&nbsp;</td>
                      </tr>
                      <tr>
                        <td background="userimages/titlebk.gif"><div align="center">
                            <table width="90%"  cellspacing="0" cellpadding="0">
                              <%if Not(rss.bof and rss.eof) then
			NumRecord=rss.recordcount
			rss.pagesize=25
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
                              <tr>
                                <td height="35"><div align="center" class="style2"><span class="style3"><%=rs("user_account")%></span> 您 的 成 绩 如 下：</div></td>
                              </tr>
                              <tr>
                                <td height="25"><div align="center">
                                    <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                      <tr>
                                        <td height="24"><div align="center"><span class="style10">课 程</span></div></td>
                                        <td height="24"><div align="center" class="style10"> 成 绩</div></td>
                                        <td height="24"><div align="center" class="style10">
                                            <div align="center">学 分</div>
                                        </div></td>
                                        <td><div align="center" class="style10">
                                            <div align="center">修读时间</div>
                                        </div></td>
                                      </tr>
                                      <%if Not(rss.bof and rss.eof) then
	rss.move (Cint(NoncePage)-1)*25,1
	for i=1 to rss.pagesize
%>
                                      <tr>
                                        <td height="24"><div align="center" class="style10">
                                            <div align="center" class="style10">
                                              <div align="center"><%=rss("course")%></div>
                                            </div>
                                        </div></td>
                                        <td height="24"><div align="center" class="style10">
                                            <div align="center"><%=rss("score")%></div>
                                        </div></td>
                                        <td height="24"><div align="center" class="style10">
                                            <div align="center"><%=rss("credit")%></div>
                                        </div></td>
                                        <td><div align="center" class="style10">
                                            <div align="center"><%=rss("term")%></div>
                                        </div></td>
                                      </tr>
                                      <%rss.movenext
     if rss.eof then exit for
	next
else
	response.write "<tr><td colspan=13><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>暂时没有成绩录入!!!</font></marquee></td></tr>"
end if	
rss.close
set rss=nothing
%>
                                      <tr>
                                        <td height="24" colspan="12"><div align="right"> <span class="style10">
                                            <input type="hidden" name="page" value="<%=NoncePage%>">
                                            <%
if NoncePage>1 then
	response.write "|<a href=user_sheet.asp?page=1>首 页</a>| |<a href=user_sheet.asp?page="&NoncePage-1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=user_sheet.asp?page="&NoncePage+1&">下一页</a>| |<a href=user_sheet.asp?page="&NumPage&">尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
&nbsp;页次：<span class="style17"><%=NoncePage%></span>/<span class="style17"><%=NumPage%></span> 共<span class="style17"><%=NumRecord%></span>条记录</span>&nbsp; </div></td>
                                      </tr>
                                    </table>
                                </div></td>
                              </tr>
                              <tr>
                                <td height="35" colspan="14"><div align="center"><a href="sheet_print.asp"><img src="userimages/print.gif" width="70" height="23" border="0" align="absmiddle"></a>&nbsp;&nbsp;&nbsp;&nbsp; <a href="user_index.asp"><img src="userimages/return.gif" width="49" height="23" border="0" align="absmiddle" ></a> </div></td>
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
<!--#include file="bottom1.asp"-->
</html>
