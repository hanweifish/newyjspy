<!--#include file="conn.asp"-->

<%
set rsn=server.createobject("adodb.recordset")
sql="select * from news order by news_time desc"
rsn.open sql,conn,1,1
%>

<script language="javascript">
window.status="��ӭ�����Ͼ���ѧ����ϵ�о���������Ϣϵͳ��"
</script>

<script language="javascript">
	function checkuser(form)
	{
		if (document.form.user_account.value=="")
		{
			alert("�������û�������");
		}
		else if (document.form.user_pwd.value=="")
		{
			alert("���������룡!");
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
<title>�о�����Ϣ����ϵͳ</title>
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
    <td rowspan="4" valign="top"><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
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
                      <td colspan="2"><div align="center">�û���:
                          <input name="user_account" type="text" class="style3" id="name" size="12">
                      </div>
                        <div align="left">                        </div></td>
                      </tr>
                    <tr>
                      <td height="38" colspan="2"><div align="center">�� &nbsp;��:
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
              <td height="120" valign="center" background="indeximages/loginbk.gif"><div align="center"><iframe src="denote.asp" name="denote" width="150" marginwidth="0" height="120" marginheight="0" align="middle" scrolling="yes" frameborder="0" allowtransparency="true" ></iframe>
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
                          <option value="javascript:void(null);" selected>----�����ѧ----</option>
                          <option value="http://www.harvard.edu/">�����ѧ</option>
                          <option value="http://www.cam.ac.uk/">���Ŵ�ѧ</option>
                          <option value="http://www.ox.ac.uk/">ţ���ѧ</option>
                          <option value="http://www.stanford.edu/">˹̹����ѧ</option>
                          <option value="http://www.yale.edu/">Ү³��ѧ</option>
                        </select>
                        </form>
                    </div></td>
                  </tr>
                  <tr>
                    <td height="50"><div align="center">
                        <form name="links">
                          <select name="links" class="style2" onChange="window.open(this.value)">
                            <option value="javascript:void(null);" selected>---ʵ��������---</option>
                            <option value="http://biophy.nju.edu.cn ">��������ʵ����</option>
                            <option value="http://pld.nju.edu.cn ">PLDʵ����</option>
                            <option value="http://x.nju.edu.cn/">�϶���С��</option>
                          </select>
                        </form>
                    </div></td>
                  </tr>
                  <tr>
                    <td height="50"><div align="center">
                        <form name="links">
                          <select name="links" class="style2" onChange="window.open(this.value)">
                            <option value="javascript:void(null);" selected>----У������----</option>
                            <option value="http://www.njbys.com/">�Ͼ���ҵ����ҵ��</option>
                            <option value="http://www.jsbys.com.cn/index.aspx">���ձ�ҵ����ҵ��</option>
                            <option value="http://www.firstjob.com.cn/">�Ϻ���ҵ����ҵ��</option>
                          </select>
                        </form>
                    </div></td>
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
              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="54" background="user/userimages/titlebk1.gif"><div align="center"><img src="indeximages/news.gif"></div></td>
                </tr>
                <tr>
                  <td height="53" background="user/userimages/titlebk2.gif">&nbsp;</td>
                </tr>
                <tr>
                  <td background="user/userimages/titlebk.gif"><div align="center">
                    <table width="90%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td colspan="6"><%if Not(rsn.bof and rsn.eof) then
			NumRecord=rsn.recordcount
			rsn.pagesize=10
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
                            <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                              <tr align="center">
                                <td width="150" height="24" align="center" valign="middle"><div align="center" class="style10"> ����ʱ�� </div></td>
                                <td width="76%" height="24"><div align="center" class="style10">ͨ &nbsp;&nbsp;&nbsp;֪ &nbsp;&nbsp;&nbsp;�� &nbsp;&nbsp;&nbsp;�� </div></td>
                              </tr>
                              <%if Not(rsn.bof and rsn.eof) then
	rsn.move (Cint(NoncePage)-1)*10,1
	for i=1 to rsn.pagesize
%>
                              <tr>
                                <td width="150" height="25" align="center" valign="middle"><div align="left" class="style10">
                                    <div align="center"><%=rsn("news_time")%></div>
                                </div></td>
                                <td height="25"><div align="left"></div>
                                    <div align="left">&nbsp;&nbsp;&nbsp;&nbsp;<a href=news_detail.asp?news_ID=<%=rsn("news_ID")%> target="_blank" class="style3"><%=rsn("news_title")%></a></div></td>
                              </tr>
                              <%rsn.movenext
if rsn.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='25'><marquee scrolldelay=120 behavior=alternate><font class='style3' color='#ff6633'>��ʱ��û��֪ͨ����!!!</font></marquee></td></tr>"
end if	
rsn.close
set rsn=nothing
%>
                              <tr>
                                <td height="24" colspan="8" align="center" valign="middle"><div align="right"><span class="style2">
                                    <input type="hidden" name="page" value="<%=NoncePage%>">
                                    <%
if NoncePage>1 then
	response.write "|<a href=newslist.asp?page=1>�� ҳ</a>| |<a href=newslist.asp?page="&NoncePage-1&">��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=newslist.asp?page="&NoncePage+1&">��һҳ</a>| |<a href=newslist.asp?page="&NumPage&">β ҳ</a>|"
else
	response.write "|��һҳ| |β ҳ|"
end if
%>
&nbsp;ҳ�Σ�<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC" class="style2"><%=NumPage%></font> ��<font color="#0033CC" class="style3"><%=NumRecord%></font>����¼</span>&nbsp; </div></td>
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
    <td height="10" valign="top">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top"><div align="right">
      <!--#include file="server.asp"-->
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
