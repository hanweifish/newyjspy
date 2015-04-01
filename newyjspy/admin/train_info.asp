<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("admin_account")="" then
Response.write"对不起，您还没有登陆，无此权限！"
Response.end
end if
%>
<%
dim admin_account
admin_account=session("admin_account")
session("admin_account")=admin_account
%>


<script language="javascript">
	function checkform()
	{
		if (document.form.user_number.value=="")
		{
			alert("请输入学号！");
		}
		else if (document.form.train_to.value=="")
		{
			alert("请输入乘车区间！");
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
                                <td><div align="center"><a href="train_infoadd.asp" target="_parent"><img src="adminimages/trainadd.gif" width="170" height="24" border="0"></a></div></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="86" background="adminimages/titlebk2.gif"><div align="center" class="style10"><img src="adminimages/train.gif" width="170" height="24"></div></td>
                      </tr>
                      <tr>
                        <td background="../user/userimages/titlebk.gif"><div align="center">
                            <table width="90%"  cellspacing="0" cellpadding="0">
                              <tr>
                                <td><div align="center">
                                    <table width="100%"  cellspacing="0" cellpadding="0">
                                      <tr>
                                        <td height="25"></td>
                                      </tr>
                                      <tr>
                                        <td><form action="train_infoadd.asp" method="post" name="form" onSubmit="return checkform(form)">
                                        </form></td>
                                      </tr>
                                    </table>
                                </div></td>
                              </tr>
                              <tr>
                                <td><div align="center">
                                    <table width="100%"  border="0" cellpadding="0" cellspacing="0">
                                      <tr>
                                        <td height="25" colspan="6"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                          <%
								 set rsS = Server.CreateObject("adodb.recordset")
								 sqlS="select user_info.user_name,user_info.user_yx,user_info.user_number,train_info.train_academy,train_info.train_info,train_info.train_ID,train_info.train_to from train_info inner join user_info on train_info.user_ID=user_info.user_ID order by train_info.train_academy"
								 rsS.open sqlS,conn,1,1
								 
								  if Not(rsS.bof and rsS.eof) then
			NumRecord=rsS.recordcount
			rsS.pagesize=10
			NumPage=rsS.Pagecount
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
                                          <tr class="style10">
                                            <td height="22"><div align="center">姓名</div></td>
                                            <td><div align="center">学号</div></td>
                                            <td><div align="center">乘车目的地</div></td>
                                            <td><div align="center">院系</div></td>
                                            <td><div align="center">附加信息</div></td>
                                            
                                            <td class="style3"><div align="center">修改</div></td>
                                            <td class="style3"><div align="center">删除</div></td>
                                          </tr>
                                         <%if Not(rsS.bof and rsS.eof) then
	rsS.move (Cint(NoncePage)-1)*10,1
	for i=1 to rsS.pagesize
%>
                                          <tr class="style10">
                                            <td height="22"><div align="center"><%=rsS("user_name")%></div></td>
                                            <td><div align="center"><%=rsS("user_number")%></div></td>
                                            <td><div align="center"><%=rsS("train_to")%></div></td>
                                            <td><div align="center"><%=rsS("user_yx")%></div></td>
                                            <td><div align="center"><%=rsS("train_info")%></div></td>
                                           
                                            <td class="style3"><div align="center"><a href=train_infomod.asp?train_ID=<%=rsS("train_ID")%> class=style3 >修改</a></div></td>
                                            <td class="style3"><div align="center"><a href=train_infodel.asp?train_ID=<%=rsS("train_ID")%> class=style3 >删除</a></div></td>
                                          </tr>
                                        
								 <%rsS.movenext
if rsS.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>没有找到任何记录!!!</font></marquee></td></tr>"
end if	
rsS.close
set rsS=nothing

								  %>
                                        </table></td>
                                      </tr>
                                      <tr>
									  <td>
									  <input type="hidden" name="page" value="<%=NoncePage%>">
                                <form name="form1" method="post" action="train_info.asp">
                                  <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                      <td valign="middle"><div align="right">
                                        <div align="right"> <span class="style2">
                                          <%
if NoncePage>1 then
	response.write "|<a href=train_info.asp?page=1>首 页</a>| |<a href=train_info.asp?page="&NoncePage-1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=train_info.asp?page="&NoncePage+1&">下一页</a>| |<a href=train_info.asp?page="&NumPage&">尾 页</a>|"
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
									  
									  
									  </td>
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
						  <td><div align="center"></div></td>
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
