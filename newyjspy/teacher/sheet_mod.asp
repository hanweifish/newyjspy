<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("admin_account")="" then
Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
Response.end
end if
%>
<%
dim user_number,course
user_number=session("user_number")
course=session("course")
session("course")=course
session("user_number")=user_number
%>
<%
dim admin_account
admin_account=session("admin_account")
%>

<%
dim sheet_id
sheet_id=trim(request("sheet_ID"))
session("sheet_ID")=sheet_id
set rs=server.createobject("adodb.recordset")
sql="select user_info.user_number,sheet.course,sheet.score,sheet.sheet_info,sheet.term,sheet.tutor from user_info inner join sheet on user_info.user_ID = sheet.user_ID where sheet_ID = "&sheet_id
rs.open sql,conn,1,1
%>

<html>
<head>
<script language="javascript">
window.status="��ӭ�����Ͼ���ѧ����ϵ�о���������Ϣϵͳ��"
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�о�����Ϣ����ϵͳ</title>
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
                                <td><div align="center"><a href="sheet_set.asp"><img src="adminimages/scorequerry.gif" width="81" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="subject_set.asp"><img src="adminimages/examcourse.gif" width="81" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="sheet_add.asp"><img src="adminimages/scoreadd.gif" width="81" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="subject_add.asp"><img src="adminimages/courseadd1.gif" width="81" height="24" border="0"></a></div></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="86" background="adminimages/titlebk2.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="45"><div align="center"><img src="adminimages/scoremag.gif" width="523" height="45"></div></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr>
                        <td background="../user/userimages/titlebk.gif"><div align="center">
                            <form name="form" method="post" action="sheet_mod1.asp">
                              <div align="center">
                                <table width="90%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                  <tr>
                                    <td height="24"><div align="right" class="style10">ѧ�ţ�&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("user_number")%> </td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">�γ̣�&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp; <%=rs("course")%> </td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">�ɼ���&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input name="score" type="text" class="style3" value="<%=rs("score")%>" ></td>
                                  </tr>
								     <tr>
                                    <td height="24"><div align="right" class="style10">�ο���ʦ��&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input name="tutor" type="text" class="style3" value="<%=rs("tutor")%>" ></td>
                                  </tr>
								  
                                  <tr>
                                    <td height="24"><div align="right"><span class="style10">�޶�ѧ�꣺&nbsp;&nbsp;</span></div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <label>
                                       <select name="year" class="style3">
                                              <%
								  if session("year")="����ѧ��"  then
								  %>
                                            <option value="��һѧ��">��һѧ��</option>
                                            <option value="�ڶ�ѧ��">�ڶ�ѧ��</option>
                                            <option value="����ѧ��" selected>����ѧ��</option>
                                              <%
								  elseif session("year")="�ڶ�ѧ��" then
								  %>
                                              <option value="��һѧ��">��һѧ��</option>
                                            <option value="�ڶ�ѧ��" selected>�ڶ�ѧ��</option>
                                            <option value="����ѧ��">����ѧ��</option>
											<%
								   else 
								  %>
                                              <option value="��һѧ��" selected>��һѧ��</option>
                                            <option value="�ڶ�ѧ��">�ڶ�ѧ��</option>
                                            <option value="����ѧ��">����ѧ��</option>
                                              <%
								 end if
								 %>
                                          </select>
                                      </label></td>
                                  </tr>
								   <tr>
                                    <td height="24"><div align="right" class="style10">�Ͽ�ѧ�ڣ�&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <select name="term" class="style3">
                                          <%
								  if session("term")="��ѧ��"  then
								  %>
                                          <option value="��ѧ��">��ѧ��</option>
                                          <option value="��ѧ��" selected>��ѧ��</option>
                                          <%
								  else
								  %>
                                          <option value="��ѧ��" selected>��ѧ��</option>
                                          <option value="��ѧ��">��ѧ��</option>
                                          <%
								 end if
								 %>
                                      </select></td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">�γ����ʣ�&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <select name="property" class="style3">
                                          <%
								  if session("property")="ѡ�޿�"  then
								  %>
                                          <option value="���޿�">���޿�</option>
                                          <option value="ѡ�޿�" selected>ѡ�޿�</option>
                                          <%
								  else
								  %>
                                          <option value="���޿�" selected>���޿�</option>
                                          <option value="ѡ�޿�">ѡ�޿�</option>
                                          <%
								 end if
								 %>
                                      </select></td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">�γ̱�ע��&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <textarea name="sheet_info" cols="36" rows="2" class="style3"><%=rs("sheet_info")%></textarea>                                    </td>
                                  </tr>
                                  <tr>
                                    <td height="35" colspan="2"><div align="center">
                                        <input type="submit" name="Submit" value="�� ��">
                                    </div></td>
                                  </tr>
                                </table>
                              </div>
                            </form>
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
                        <td width="60%" class="style3"><%=admin_account%>��</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr>
                        <td ><div align="center"> </div>
                            <div align="center"></div></td>
                        <td height="30" class="style2">���Ѿ���¼�ɹ�,��</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr>
                        <td><div align="center"></div></td>
                        <td height="30" class="style2">������ά����վ!</td>
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