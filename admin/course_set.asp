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

<%
set rs=server.createobject("adodb.recordset")
sql="select * from course order by course_term "
rs.open sql,conn,1,1
set rsc=server.createobject("adodb.recordset")
sql="select * from course_set"
rsc.open sql,conn,1,3
%>

<%
if trim(request("select"))<>"" then
rsc("select")=trim(request("select"))
response.write"<script>alert('����ѡ��״̬�ɹ���')</script>"
end if
%>

<%
if trim(request("enddate"))<>"" then
rsc("enddate")=trim(request("enddate"))
response.write"<script>alert('����ѡ�ν������ڳɹ���\n     "&rsc("enddate")&"')</script>"
end if
if trim(request("startdate"))<>"" then
rsc("startdate")=trim(request("startdate"))
response.write"<script>alert('����ѡ����ʼ���ڳɹ���\n     "&rsc("startdate")&"')</script>"
end if
%>

<%
rsc.update
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
                        <td background="../user/userimages/titlebk.gif"><div align="center">
                            <table width="90%"  border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <form name="form1" method="post" action="">
                                  <td height="30" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ѡ�ο�ʼʱ���趨&nbsp;��&nbsp;&nbsp;
                                      <input name="startdate" type="text" class="style3" id="startdate" size="20" maxlength="20" value="<%=rsc("startdate")%>">
&nbsp;<span class="style3">���밴��ʽ���룺2005-05-18�����س�ȷ��</span></td>
                                </form>
                              </tr>
                              <tr>
                                <form name="form1" method="post" action="">
                                  <td height="30" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ѡ�ν�ֹʱ���趨&nbsp;��&nbsp;&nbsp;
                                      <input name="enddate" type="text" class="style3" id="enddate" size="20" maxlength="20" value="<%=rsc("enddate")%>">
&nbsp;<span class="style3">���밴��ʽ���룺2005-05-18�����س�ȷ��</span></td>
                                </form>
                              </tr>
                              <tr>
                                <td height="30">
                                  <div align="center">
                                    <input name="select" type="button" class="style8" id="select" onclick ="window.location.href='course_set.asp?select=yes'" value="��ʼѡ�ΰ�ť">
                                  </div></td>
                                <td><div align="center">
                                    <input name="select" type="button" class="style3" id="select" onclick ="window.location.href='course_set.asp?select=no'" value="ǿ����ֹѡ��">
                                  </div></td>
                              </tr>
                              <tr>
                                <td height="25" colspan="2">                                  <div align="center"><span class="style2">ѡ �� �� �� �� ��</span> </div></td>
                              </tr>
                              <tr>
                                <td colspan="2"><%       if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=30
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
                                    <table width="100%" height="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                      <tr>
                                        <td height="24"><div align="center" class="style10">�γ�����</div></td>
                                        <td width="80" height="24"><div align="center" class="style10">�γ̱��</div></td>
                                        <td width="80" height="24"><div align="center" class="style10">������ʦ</div></td>
                                        <td width="50" height="24"><div align="center" class="style10">ѧ��</div></td>
                                        <td width="80" height="24"><div align="center" class="style10">�Ƿ����ѡ��</div></td>
                                        <td width="50" height="24"><div align="center" class="style10">�޸�</div></td>
                                        <td width="50" height="24"><div align="center" class="style10">ɾ��</div></td>
                                      </tr>
                                      <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*30,1
	for i=1 to rs.pagesize
%>
                                      <tr>
                                        <td height="24"><div align="center" class="style10"><%=rs("course_name")%></div></td>
                                        <td width="80" height="24"><div align="left" class="style10">&nbsp;<%=rs("course_number")%></div></td>
                                        <td width="80" height="24"><div align="center" class="style10"><%=rs("course_tutor")%></div></td>
                                        <td width="50" height="24"><div align="center" class="style10"><%=rs("course_credit")%></div></td>
                                        <td width="80" height="24"><div align="center" class="style10">
                                            <%
								  if rs("course_term") = "1" then
								  %>
                              ��
                              <%
								  else
								  %>
                              ��
                              <%
								  end if
								  %>
                                        </div></td>
                                        <td width="50" height="24"><div align="center"><a href=course_mod.asp?course_ID=<%=rs("course_ID")%>><font color="#ff6633">�޸�</font></a></div></td>
                                        <td width="50" height="24"><div align="center"><a href=course_del.asp?course_ID=<%=rs("course_ID")%>><font color="#ff6633">ɾ��</font></a></div></td>
                                      </tr>
                                      <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>û���ҵ��κμ�¼!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
                                      <tr>
                                        <td height="24" colspan="8"><div align="right"> <span class="style10">
                                            <input type="hidden" name="page" value="<%=NoncePage%>">
                                            <%
if NoncePage>1 then
	response.write "|<a href=course_set.asp?page=1>�� ҳ</a>| |<a href=course_set.asp?page="&NoncePage-1&">��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=course_set.asp?page="&NoncePage+1&">��һҳ</a>| |<a href=course_set.asp?page="&NumPage&">β ҳ</a>|"
else
	response.write "|��һҳ| |β ҳ|"
end if
%>
&nbsp;ҳ�Σ�<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> ��<font color="#0033CC" class="style3"><%=NumRecord%></font>����¼</span>&nbsp; </div></td>
                                      </tr>
                                  </table></td>
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
<%
rsc.close
set rsc=nothing
%>
<!--#include file="bottom1.asp"-->
</html>