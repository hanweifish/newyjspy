<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("admin_account")="" or session("user_group")="" then
Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
Response.end
end if
%>

<%
dim admin_account
admin_account=session("admin_account")
%>

<script language="javascript">
	function checkcourse(form)
	{
		if (document.form.user_number.value=="")
		{
			alert("������ѧ��ѧ�ţ�");
		}
		else if (document.form.score.value=="")
		{
			alert("������ɼ���");
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
<!--#include file="top2.asp"-->
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
                        <td valign="top" background="../user/userimages/titlebk.gif"><div align="center">
                          <table width="90%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><div align="center">
                                <form name="form" method="post" action="sheet_add1.asp" onSubmit="return checkcourse(form)">
                                  <div align="center">
                                    <table width="90%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                      <tr>
                                        <td height="24"><div align="right" class="style10">�γ̣�&nbsp;&nbsp;</div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp;
                                            <select name="course_ID" class="style3">
                                              <%
												set rs1=server.createobject("adodb.recordset")
												sql1="select * from subject order by course"
												rs1.open sql1,conn,1,1
												%>
                                              <%do while not rs1.eof%>
                                              <%
												if session("course")=rs1("course") then
												%>
                                              <option value="<%=rs1("course_ID")%>" selected><%=rs1("course")&" ("&rs1("tutor")&"-"&rs1("credit")&"ѧ��"&"-"&rs1("term")&"-"&rs1("teachway")&")"%></option>
                                              <%
									else
									%>
                                              <option value="<%=rs1("course_ID")%>"><%=rs1("course")&" ("&rs1("tutor")&"-"&rs1("credit")&"ѧ��"&"-"&rs1("term")&"-"&rs1("teachway")&")"%></option>
                                              <%
									end if
									%>
                                              <%rs1.movenext%>
                                              <%loop%>
                                              <%rs1.close%>
                                          </select></td>
                                      </tr>
                                      <tr>
                                        <td height="24"><div align="right" class="style10">ѧ�ţ�&nbsp;&nbsp;</div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp;
                                            <%
									if session("user_number") <> "" then
									%>
                                            <input name="user_number" type="text" class="style3" value="<%=session("user_number")%>" >
                                            <%
									else
									%>
                                            <input name="user_number" type="text" class="style3" value="" >
                                            <%
									end if
									%>                                        </td>
                                      </tr>
									  <tr>
                                        <td height="24"><div align="right" class="style10">�ο���ʦ��&nbsp;&nbsp;</div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp;
										<%
									if session("tutor") <> "" then
									%>
                                            <input name="tutor" type="text" class="style3" value="<%=session("tutor")%>" >
                                            <%
									else
									%>
                                            <input name="tutor" type="text" class="style3" value="" >
                                            <%
									end if
									%>                                        </td>
                                      </tr>
									  <tr>
                                        <td height="24"><div align="right"><span class="style10">�޶�ѧ�꣺&nbsp;&nbsp;</span></div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp; <label>
                                          <select name="term" class="style3" id="term">
                                             <%
								  if session("term")="��ѧ��"  then
								  %>
											<option value="��ѧ��" selected>��ѧ��</option>
                                            <option value="��ѧ��">��ѧ��</option>
									<%
								  else
								  %>
                                              <option value="���޿�">���޿�</option>
                                              <option value="ѡ�޿�" selected>ѡ�޿�</option>
                                              <%
								 end if
								 %>
                                          </select>
                                        </label></td>
                                      </tr>
                                      <tr>
                                        <td height="24"><div align="right" class="style10">�ɼ���&nbsp;&nbsp;</div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp;
                                            <input name="score" type="text" class="style3" ></td>
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
                                        <td height="24"><div align="right"><span class="style10">�޶�ѧ�꣺&nbsp;&nbsp;</span></div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp; <label>
                                          <select name="year" class="style3" id="year">
                                            <option value="��һѧ��" selected>��һѧ��</option>
                                            <option value="�ڶ�ѧ��">�ڶ�ѧ��</option>
                                            <option value="����ѧ��">����ѧ��</option>
                                          </select>
                                        </label></td>
                                      </tr>
                                      <tr>
                                        <td height="24"><div align="right" class="style10">�γ̱�ע��&nbsp;&nbsp;</div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp;
                                            <textarea name="sheet_info" cols="36" rows="2" class="style3">��</textarea>                                        </td>
                                      </tr>
                                      <tr>
                                        <td height="25" colspan="2"><div align="center">
                                            <input type="submit" name="Submit" value="�� ��" >
                                        </div></td>
                                      </tr>
                                    </table>
                                  </div>
                                </form>
                              </div></td>
                            </tr>
                            <tr>
                              <td><div align="center">
                                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                 <%
								 set rsS = Server.CreateObject("adodb.recordset")
								 sqlS="select top 6 user_info.user_name,user_info.user_number,sheet.score,sheet.course,sheet.property,sheet.sheet_time,sheet.sheet_ID from sheet inner join user_info on sheet.user_ID=user_info.user_ID order by sheet_time desc"
								 rsS.open sqlS,conn,1,1
								 %>
								  <tr class="style10">
                                    <td height="22"><div align="center">����</div></td>
                                    <td><div align="center">ѧ��</div></td>
                                    <td><div align="center">�γ�</div></td>
                                    <td><div align="center">�ɼ�</div></td>
                                    <td><div align="center">�γ�����</div></td>
                                    <td class="style3"><div align="center">�޸�</div></td>
                                    <td class="style3"><div align="center">ɾ��</div></td>
								  </tr>
								 <%
								 if not (rsS.eof and rsS.bof) then
								 for i=0 to rsS.recordcount
								 %>
								  <tr class="style10">
                                    <td height="22"><div align="center"><%=rsS("user_name")%></div></td>
                                    <td><div align="center"><%=rsS("user_number")%></div></td>
                                    <td><div align="center"><%=rsS("course")%></div></td>
                                    <td><div align="center"><%=rsS("score")%></div></td>
                                    <td><div align="center"><%=rsS("property")%></div></td>
                                    <td class="style3"><div align="center"><a href=sheet_mod.asp?sheet_ID=<%=rsS("sheet_ID")%> class=style3 >�޸�</a></div></td>
                                    <td class="style3"><div align="center"><a href=sheet_del.asp?sheet_ID=<%=rsS("sheet_ID")%> class=style3 >ɾ��</a></div></td>
								  </tr>
								  <%
								  rsS.movenext
								  if rsS.eof then exit for 
								  next
								  end if
								  rsS.close
								  set rsS=nothing
								  %>
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
