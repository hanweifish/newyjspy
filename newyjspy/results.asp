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

<%
dim user_name
user_name=trim(request("user_name"))
set rs1=server.createobject("adodb.recordset")
sql="select * from user_info where user_name='"&user_name&"'"
rs1.open sql,conn,1,1
%>

<script language="javascript">
window.status="��ӭ�����Ͼ���ѧ����ϵ�о���������Ϣϵͳ��"
</script>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="����ϵ�о�����Ϣ�����Ͼ���ѧ">
<meta name="description" content="�Ͼ���ѧ����ϵ�о���ע��ϵͳ������ѡ�Σ���Ҫ��Ϣ����">
<title>�о���Ժ��Ϣ����ϵͳ</title>
<link href="style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
-->
</style>
<!--#include file="top.asp"-->
<table width="800"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="187" height="10"></td>
    <td rowspan="7" valign="top"><div align="right">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="22"></td>
        </tr>
        <tr>
          <td valign="top"><div align="right">
            <table width="603"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="406" height="55" background="indeximages/newstudent.jpg">&nbsp;</td>
                <td width="6" background="indeximages/midLinkTop.gif">&nbsp;</td>
                <td background="indeximages/bulletin.gif">&nbsp;</td>
              </tr>
              <tr>
                <td height="25" background="indeximages/noticebktop.gif">&nbsp;</td>
                <td background="indeximages/midLinkTop.gif">&nbsp;</td>
                <td background="indeximages/bulletinbktop.gif">&nbsp;</td>
              </tr>
              <tr>
                <td><table width="100%" height="180" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td background="indeximages/noticeBk.gif">
					  <div align="center">
					 
					    <p><span class="style2">��������:</span>					      <font class="style3"><%=user_name%></font></p>
					    <p><span class="style2">����ѧ��:</span><font class="style3">						<%=rs1("user_number")%></font><br>
					      <br>
					        </p>
					    <div align="center"><br><br>
					      <a href="search_number.asp">����</a></div>
			
					  
					  </div></td>
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
                <td height="60" valign="top" background="indeximages/noticeBt.gif"><div align="center"><a href="noticelist.asp" class="style3" target="_blank"></a></div></td>
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
    <td height="45" background="indeximages/stulogin.gif">&nbsp;</td>
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
                        <div align="left"> </div></td>
                  </tr>
                  <tr>
                    <td height="38" colspan="2"><div align="center">�� &nbsp;��:
                            <input name="user_pwd" type="password" class="style3" id="pwd" size="12">
                      </div>
                        <div align="left"> </div></td>
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
          <td height="120" valign="center" background="indeximages/loginbk.gif"><div align="center">
              <iframe src="denote.asp" name="denote" width="140" marginwidth="0" height="110" marginheight="0" align="middle" scrolling="yes" frameborder="0" allowtransparency="true" ></iframe>
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
                          <option value="javascript:void(null);" selected>----У������----</option>
                          <option value="http://www.nju.edu.cn/">�Ͼ���ѧ</option>
                          <option value="http://lily.nju.edu.cn/">�ϴ�С�ٺ�</option>
                          <option value="http://grawww.nju.edu.cn/">�о���Ժ</option>
                          <option value="http://physics.nju.edu.cn/">����ѧϵ</option>
                          <option value="http://job.nju.edu.cn/">��ҵָ������</option>
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
<%
rs1.close
set rs1=nothing
%>
<!--#include file="bottom.asp"-->
</html>
