<!--#include file="conn.asp"-->

<%
set rsn=server.createobject("adodb.recordset")
sql="select * from tongzhi order by tongzhi_time desc"
rsn.open sql,conn,1,1
%>

<script language="javascript">
window.status="欢迎访问南京大学物理系研究生管理信息系统！"
</script>

<script language="javascript">
	function checkuser(form)
	{
		if (document.form.user_account.value=="")
		{
			alert("请输入用户名！！");
		}
		else if (document.form.user_pwd.value=="")
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
<style type="text/css">
<!--
.style10 {font-size: 12px;
	color: #004080;
}
-->
</style>
<style type="text/css">
<!--
.style12 {color: #FF0000}
-->
</style>
<!--#include file="top.asp"-->
<table width="800"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="187" height="12"></td>
    <td></td>
  </tr>
  <tr>
    <td rowspan="3" valign="top"><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="47" background="indeximages/stulogin.gif">&nbsp;</td>
      </tr>
      <tr>
        <td><div align="center">
          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="132"><div align="center">
                <form action="user/user_check.asp" method="post" name="form" id="form">
                  <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" background="indeximages/loginbk.gif">
                    <tr>
                      <td colspan="2"><div align="center">用户名:
                          <input name="user_account" type="text" class="style3" id="name" size="12">
                      </div>
                        <div align="left">                        </div></td>
                      </tr>
                    <tr>
                      <td height="38" colspan="2"><div align="center">密 &nbsp;码:
                          <input name="user_pwd" type="password" class="style3" id="pwd" size="12"> 
                        </div>
                        <div align="left">                        </div></td>
                      </tr>
                    <tr>
                      <td width="50%"><div align="right"><img src="indeximages/login.gif" width="49" height="23" border="0" style='cursor:hand' onMouseDown="checkuser(form)">&nbsp;</div></td>
                      <td width="50%"><div align="left">&nbsp;<a href="javascript:void(null)"><img src="indeximages/register.gif" width="49" height="23" border="0"></a></div></td>
                    </tr>
                    <tr>
                      <td height="10"><div align="center"></div></td>
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                </form>
                </div></td>
            </tr>
            <tr>
              <td height="6" background="indeximages/loginbk.gif"><div align="center"><img src="indeximages/loginbar.gif" width="129" height="2"></div></td>
            </tr>
            <tr>
              <td height="200" valign="center" background="indeximages/loginbk.gif"><div align="center"><iframe src="denote.asp" name="denote" width="150" marginwidth="0" height="200" marginheight="0" align="middle" scrolling="yes" frameborder="0" allowtransparency="true" ></iframe>
              </div></td>
            </tr>
          </table>
        </div></td>
      </tr>
      <tr>
        <td height="77" background="indeximages/links.gif">&nbsp;</td>
      </tr>
      <tr>
        <td background="indeximages/loginbk.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="50"><div align="center"><a href="http://job.nju.edu.cn/">南大就业指导中心</a></div></td>
          </tr>
          <tr>
            <td height="40"><div align="center"><a href="http://www.njbys.com/">南京毕业生就业网</a></div></td>
          </tr>
          <tr>
            <td height="40"><div align="center"><a href="http://www.jsbys.com.cn/index.aspx">江苏毕业生就业网</a></div></td>
          </tr>
          <tr>
            <td height="40"><div align="center"><a href="http://www.js.lm.gov.cn/">中国劳动力市场网</a></div></td>
          </tr>
          <tr>
            <td height="40"><div align="center"><a href="http://www.firstjob.com.cn/">上海毕业生就业网</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center">&nbsp;&nbsp;&nbsp;&nbsp;<a href="links.asp" class="style3">&gt;&gt;&gt; More</a></div></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="34" background="indeximages/loginbottom.gif">&nbsp;</td>
      </tr>
    </table></td>
    <td valign="top"><div align="right">
      <table width="603"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="406" height="10">&nbsp;</td>
          <td width="406" background="../indeximages/midLinkTop.gif">&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td colspan="3"><div align="center">
              <table width="100%" height="700"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="54" background="user/userimages/titlebk1.gif"><div align="center"><img src="user/userimages/notice.gif" width="523" height="45"></div></td>
                </tr>
                <tr>
                  <td height="53" background="user/userimages/titlebk2.gif">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" background="user/userimages/titlebk.gif"><div align="center">
                    <table width="90%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td colspan="6" valign="top">
	<%if Not(rsn.bof and rsn.eof) then
			NumRecord=rsn.recordcount
			rsn.pagesize=26
			NumPage=rsn.Pagecount
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
                            <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" class="thin">
                              <tr align="center">
                                <td width="76%" height="20" bgcolor="#F4F7FB"><div align="left" class="style10">&nbsp;&nbsp;发布时间&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;通 &nbsp;&nbsp;&nbsp;知 &nbsp;&nbsp;&nbsp;标 &nbsp;&nbsp;&nbsp;题 </div></td>
                              </tr>
<%if (rsn.bof or rsn.eof) then
	response.write "<tr><td colspan=13 height='25'><marquee scrolldelay=120 behavior=alternate><font class='style3' color='#ff6633'>暂时还没有通知发布!!!</font></marquee></td></tr>"
  else
	rsn.move (Cint(NoncePage)-1)*26,1
	for i=1 to rsn.pagesize
%>
                              <tr>
							  <%
							  if i mod 2 = 1 then 
							  %>
                                <td height="20"><div align="left">&nbsp;&nbsp;<%=rsn("tongzhi_time")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=tongzhi_detail.asp?tongzhi_ID=<%=rsn("tongzhi_ID")%>  target="_blank" class="style3"><font color="#000000"><%=rsn("tongzhi_title")%></font></a></div></td>
							<%
								else
							%>
                                <td height="20" bgcolor="#F4F7FB">
                                    <div align="left">&nbsp;&nbsp;<%=rsn("tongzhi_time")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=tongzhi_detail.asp?tongzhi_ID=<%=rsn("tongzhi_ID")%>  target="_blank" class="style3"><font color="#000000"><%=rsn("tongzhi_title")%></font></a></div></td>
							<%
								end if
							%>
                              </tr>
<%rsn.movenext
if rsn.eof then exit for
	next
end if	
rsn.close
set rsn=nothing
%>
                              <tr>
                                <td height="20" colspan="8" align="center" valign="middle"><div align="right"><span class="style8">
                                    <input type="hidden" name="page" value="<%=NoncePage%>">
                                    <%
if NoncePage>1 then
	response.write "|<a href=tongzhi.asp?page=1>首 页</a>| |<a href=tongzhi.asp?page="&NoncePage-1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=tongzhi.asp?page="&NoncePage+1&">下一页</a>| |<a href=tongzhi.asp?page="&NumPage&">尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
&nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC" class="style2"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录</span>&nbsp; </div></td>
                              </tr>
                          </table></td>
                      </tr>
                    </table>
                  </div></td>
                </tr>
                <tr>
                  <td height="34" background="user/userimages/titlebk3.gif">&nbsp;</td>
                </tr>
              </table>
          </div></td>
        </tr>
      </table>
    </div></td>
  </tr>
  <tr>
    <td valign="top"><div align="right">
    </div></td>
  </tr>
  <tr>
    <td rowspan="2" valign="top">
      <div align="right">	    </div></td>
  </tr>
  <tr>
    <td height="12"></td>
    </tr>
</table>
<!--#include file="bottom.asp"-->
</html>
