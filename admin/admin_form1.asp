<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("admin_account")="" or session("user_group")="subadmin" then
Response.write"�Բ�������û�е�½�����޴�Ȩ�ޣ�"
Response.end
end if
%>

<%
dim admin_account
admin_account=session("admin_account")
user_number=trim(request("studentno"))
set rs1=server.createobject("adodb.recordset")
sql1="select * from user_info where user_number='"&user_number&"'"
rs1.open sql1,conn,1,1

set rs=server.createobject("adodb.recordset")
sql="select * from user_info order by user_id "
rs.open sql,conn,1,1
rs.movefirst

%>



<script language="javascript">
	function checkform(form3)
	{
		if (document.form3.studentno.value=="")
		{
			alert("������ѧ��ѧ�ţ�");
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
<%
	if session("popup") = "" then
	set rsnew = server.createobject("adodb.recordset")
	newsql = "select * from guestbook where admin_read = false "
	rsnew.open newsql,conn,1,1
	if not(rsnew.eof or rsnew.bof) then
	Response.write "<script> alert('���������ԣ�')</script> "
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
                <td height="54" background="adminimages/titlebk1.gif"><div align="center">
                    <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td><div align="center"><a href="ynumber_set.asp"></a><a href="info_search.asp"><img src="adminimages/stuquerry.gif" width="134" height="24" border="0"></a></div></td>
                        <td><div align="center"><a href="info_add.asp"><img src="adminimages/infoadd.gif" width="134" height="24" border="0"></a></div></td>
                        <td><div align="center"><a href="./jiangcheng.asp" target="_blank" class="style3">ѧ��������Ϣ����</a></div></td>
                        <td>&nbsp;</td>
                      </tr>
                    </table>
                </div></td>
              </tr>
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
                        <div align="center" class="style10">
                          <form name="form3" method="post" action="admin_form.asp">
                            ���뿪ʼ��ʾ��ѧ�ţ�<input name="studentno" type="text" class="style3" id="studentno" size="10">
&nbsp;&nbsp;                            <img src="../user/userimages/search.gif" width="46" height="23" align="absbottom" style="cursor:hand" onMouseDown="javascript:checkform(form3)">
						
                          </form>
                          <br>                   
						  <%
if not rs1.eof then
%>
                                      <%
do while not rs1.eof
%>
						  <%
		dim studentid,studentid1,studentid2
		studentid1=cint(trim(rs1("user_id")))
		studentid2=cint(trim(rs("user_id")))
studentid=studentid1-studentid2
if studentid>=350 then 
  studentid=studentid-2
  if  studentid>=360 then
  studentid=studentid-1
    if  studentid>=375 then
	   studentid=studentid-1
     end if
  end if
end if
						  %>
						  <%      
			
			if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount-studentid
			rs.pagesize=10
			NumPage=INT(NumRecord/rs.pagesize)+1
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
                              <td width="80" height="24"><div align="center" class="style10">����</div></td>
                                  <td width="80" height="24"><div align="center" class="style10">ѧ��</div></td>
                                  <td height="24"><div align="center" class="style10">רҵ</div></td>
                                  <td width="50" height="24"><div align="center" class="style10">�Ա�</div></td>
                                  <td width="50" height="24"><div align="center" class="style10">�ɼ���</div></td>
                                  <td width="80" height="24"><div align="center" class="style10"><span class="style3">��ҵ��Ϣ</span></div></td>
                                  <td width="50" height="24"><div align="center" class="style10">�޸�</div></td>
                                  <td width="50" height="24"><div align="center" class="style10">ɾ��</div></td>
                                </tr>
                            <%
	if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*10+studentid,1
	for i=1 to rs.pagesize
%>
                            <tr>
                              <td width="80" height="24"><div align="center" class="style10"><a href=name_search.asp?user_name=<%=rs("user_name")%> class="style3"><%=rs("user_name")%></a></div></td>
                                  <td width="80" height="24"><div align="center" class="style10"><%=rs("user_number")%></div></td>
                                  <td height="24"><div align="center" class="style10"><%=rs("user_major")%></div></td>
                                  <td width="50" height="24"><div align="center" class="style10"><%=rs("user_sex")%></div></td>
                                  <td width="50" height="24"><div align="center" class="style10"><img src="../images/sheet.png" alt="�ɼ���" width="25" height="24" style="cursor:hand" onClick="window.navigate('user_sheet.asp?user_account=<%=rs("user_account")%>')"></div></td>
								  <%
									set rsBiyeInfo = Server.CreateObject("Adodb.Recordset")
									sql_rsBiyeInfo = "Select * from biyeInfo where user_number = '"&rs("user_number")&"'"
									rsBiyeInfo.Open sql_rsBiyeInfo,conn,1,1
									
									if rsBiyeInfo.RecordCount = 0 then
								%>
                              <td width="80" height="24"><div align="center"><a href="AddbiyeInfo.asp?page=<%=NoncePage%>&user_number=<%=rs("user_number")%>" target="_self" class="style3">���</a></div></td>
								  <%
									elseif rsBiyeInfo("bylw") = "����" then
								%>
                              <td width="80" height="24"><div align="center"><a href="AddbiyeInfo.asp?page=<%=NoncePage%>&user_number=<%=rs("user_number")%>" target="_self" class="style3">���</a></div></td>
								  <%
									else
								%>
                              <td width="80" height="24"><div align="center"><a href="ViewbiyeInfo.asp?page=<%=NoncePage%>&user_number=<%=rs("user_number")%>" target="_self" class="style3">�鿴</a></div></td>
								  <%
									end if
									rsBiyeInfo.Close
								%>
                              
                              <td width="50" height="24"><div align="center" class="style3"><a href=info_set.asp?NoncePage=<%=NoncePage%>&user_ID=<%=rs("user_ID")%>><font color="#ff6633">�޸�</font></a></div></td>
                                  <td width="50" height="24"><div align="center" class="style3"><a href=info_del.asp?NoncePage=<%=NoncePage%>&user_ID=<%=rs("user_ID")%>><font color="#ff6633">ɾ��</font></a></div></td>
                                </tr>
								
                           
                            <tr>
                              <td height="40" colspan="8"><div align="center">
                                <input type="hidden" name="page" value="<%=NoncePage%>">
                
                                  <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                      <td valign="middle">&nbsp;</td>
                                    </tr>
                                    </table>
                                          </form>
                                  </div></td>
                                </tr>
                            </table>
 <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>û���ҵ��κμ�¼!!!</font></marquee></td></tr>"
	end if
%>
 <%
rs1.movenext
loop
else
response.write "<tr><td height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>û���ҵ��κμ�¼!!!</font></marquee></td></tr>"                                                             
end if
%>
<%
rs1.close
set rs1=nothing
rs.close
set rs=nothing
%>
                        </div>
						</td>
						</tr>
						                                     
                    </table>
                </div></td>
              </tr>
              <tr>
                <td height="40" background="adminimages/titlebk3.gif">&nbsp;</td>
              </tr>
            </table>
          </div></td>
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
