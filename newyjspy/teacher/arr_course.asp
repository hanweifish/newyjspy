<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("admin_account")=""then
Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
Response.end
end if
%>

<%
dim admin_account
admin_account=session("admin_account")
dim admin_academy
set rs2=server.createobject("adodb.recordset")
sql2="select * from teacher_info where admin_account='"&admin_account&"'"
rs2.open sql2,conn,1,1
admin_academy=rs2("admin_academy")
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
                               
                                
                                <td><div align="center" class="style10">�ſ�ϵͳ</div></td>
                               
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="86" background="adminimages/titlebk2.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="45"><div align="center"><img src="adminimages/arr_course.gif" width="523" height="45"></div></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr>
                        <td valign="top" background="../user/userimages/titlebk.gif"><div align="center">
                          <table width="90%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><div align="center">
                                <form name="form" method="post" action="arr_courseadd.asp" onSubmit="return checkcourse(form)">
                                  <div align="center">
                                    <table width="90%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                      <tr>
                                        <td height="24"><div align="right" class="style10">�γ̣�&nbsp;&nbsp;</div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp;
                                            <select name="course_ID" class="style3">
                                              <%
												set rs1=server.createobject("adodb.recordset")
												sql1="select * from subject where course_academy='"&admin_academy&"' order by course "
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
                                        <td height="24"><div align="right" class="style10">�������ڣ�&nbsp;&nbsp;</div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp;
                                            <select name="theweek" class="style3">
											<option value="1" selected="selected">����һ</option>
											<option value="2">���ڶ�</option>
											<option value="3">������</option>
											<option value="4">������</option>
											<option value="5">������</option>
											</select></td>
                                      </tr>
									  <tr>
                                        <td height="24"><div align="right" class="style10">���ſ�ʱ��&nbsp;&nbsp;</div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp;
<input type="radio" name="coursetime" value="1">1,2�ڿ�
<input type="radio" name="coursetime" value="2">3,4�ڿ�
<input type="radio" name="coursetime" value="3">5,6�ڿ�
<input type="radio" name="coursetime" value="4">7,8�ڿ�

</td>
                                      </tr>
                                      <tr>
                                        <td height="24"><div align="right" class="style10">�Ƿ�˫�ܣ�&nbsp;&nbsp;</div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp;
                                            <select name="doubleweek" class="style3">
                                             
                                              <option value="no" selected>��</option>
                                              <option value="yes">��</option>
                                   
                                          </select></td>
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
							 <tr><td height="30"></td></tr>
							<tr>
							<td align="center"><table width="100%"  border="1" cellspacing="0" cellpadding="0">
                                          <%
								 set rsS = Server.CreateObject("adodb.recordset")
								 sqlS="select  arr_course.arr_courseid, arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy,subject.period from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"'  order by arr_course.arr_courseid desc"
								 rsS.open sqlS,conn,1,1
								 %>
                                          <tr class="style10">
                                            <td height="22"><div align="center">�γ�</div></td>
                                            <td><div align="center">ѧʱ</div></td>
                                            <td><div align="center">�Ͽ�ʱ��</div></td>
                                            
                                            <td class="style3"><div align="center">ɾ��</div></td>
                                          </tr>
                                          <%
								 if not (rsS.eof and rsS.bof) then
								 for i=0 to rsS.recordcount
								 %>
                                          <tr class="style10">
                                            <td height="22" width="150"><div align="left"><%=rsS("course")%></div></td>
                                            <td><div align="center"><%=rsS("period")%></div></td>
                                            <td><div align="center"><%=rsS("theweek")%>,<%=rsS("coursetime")%></div></td>
                                            
                                        
                                            <td class="style3"><div align="center"><a href=arr_coursedel.asp?arr_courseid=<%=rsS("arr_courseid")%> class=style3 >ɾ��</a></div></td>
                                          </tr>
                                          <%
								  rsS.movenext
								  if rsS.eof then exit for 
								  next
								  end if
								  rsS.close
								  set rsS=nothing
								  %>
                                        </table></td>
							</tr>
							 <tr><td height="30"></td></tr>
                            <tr>
                              <td><div align="center">
                                <table width="100%"  border="1" cellspacing="0" cellpadding="0">
                                
								  <tr class="style10">
                                     <td height="22" width=50><div align="center"></div></td>
									<td height="22"><div align="center">����һ</div></td>
                                    <td><div align="center">���ڶ�</div></td>
                                    <td><div align="center">������</div></td>
									<td><div align="center">������</div></td>
                                    <td><div align="center">������</div></td>
                                   
								  </tr>
								 
								  <tr class="style10">
								   <td height="22"><div align="center">1.2��</div></td>
								   
                                    <td height="22"><div align="center">
									<%
								 set rsS11 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='1' and arr_course.theweek='1' order by arr_course.theweek"
								 rsS11.open sqlS,conn,1,1
								
								 if not (rsS11.eof and rsS11.bof) then
								 for i=0 to rsS11.recordcount
								 %>
								<%=rsS11("course")%>
                                  <%
								  rsS11.movenext
								  if rsS11.eof then exit for 
								  next
								  end if
								  rsS11.close
								  set rsS11=nothing
								  %>
                                   </div></td> 
								  
							
							                                    <td height="22"><div align="center">
									<%
								 set rsS21 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='1' and arr_course.theweek='2' order by arr_course.theweek"
								 rsS21.open sqlS,conn,1,1
								
								 if not (rsS21.eof and rsS21.bof) then
								 for i=0 to rsS21.recordcount
								 %>
								<%=rsS21("course")%>
                                  <%
								  rsS21.movenext
								  if rsS21.eof then exit for 
								  next
								  end if
								  rsS21.close
								  set rsS21=nothing
								  %>
                                   </div></td> 
								   
								  <td height="22"><div align="center">
									<%
								 set rsS31 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='1' and arr_course.theweek='3' order by arr_course.theweek"
								 rsS31.open sqlS,conn,1,1
								
								 if not (rsS31.eof and rsS31.bof) then
								 for i=0 to rsS31.recordcount
								 %>
								<%=rsS31("course")%>
                                  <%
								  rsS31.movenext
								  if rsS31.eof then exit for 
								  next
								  end if
								  rsS31.close
								  set rsS31=nothing
								  %>
                                   </div></td>  
								   
								   <td height="22"><div align="center">
									<%
								 set rsS41 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='1' and arr_course.theweek='4' order by arr_course.theweek"
								 rsS41.open sqlS,conn,1,1
								
								 if not (rsS41.eof and rsS41.bof) then
								 for i=0 to rsS41.recordcount
								 %>
								<%=rsS41("course")%>
                                  <%
								  rsS41.movenext
								  if rsS41.eof then exit for 
								  next
								  end if
								  rsS41.close
								  set rsS41=nothing
								  %>
                                   </div></td>   
								    <td height="22"><div align="center">
									<%
								 set rsS51 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='1' and arr_course.theweek='5' order by arr_course.theweek"
								 rsS51.open sqlS,conn,1,1
								
								 if not (rsS51.eof and rsS51.bof) then
								 for i=0 to rsS51.recordcount
								 %>
								<%=rsS51("course")%>
                                  <%
								  rsS51.movenext
								  if rsS51.eof then exit for 
								  next
								  end if
								  rsS51.close
								  set rsS51=nothing
								  %>
                                   </div></td>  
								  </tr>
							 <tr class="style10">
								   <td height="22"><div align="center">3.4��</div></td>
								   
                                    <td height="22"><div align="center">
									<%
								 set rsS21 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='2' and arr_course.theweek='1' order by arr_course.theweek"
								 rsS21.open sqlS,conn,1,1
								
								 if not (rsS21.eof and rsS21.bof) then
								 for i=0 to rsS21.recordcount
								 %>
								<%=rsS21("course")%>
                                  <%
								  rsS21.movenext
								  if rsS21.eof then exit for 
								  next
								  end if
								  rsS21.close
								  set rsS21=nothing
								  %>
                                   </div></td> 
								  
							
							                                    <td height="22"><div align="center">
									<%
								 set rsS22= Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='2' and arr_course.theweek='2' order by arr_course.theweek"
								 rsS22.open sqlS,conn,1,1
								
								 if not (rsS22.eof and rsS22.bof) then
								 for i=0 to rsS22.recordcount
								 %>
								<%=rsS22("course")%>
                                  <%
								  rsS22.movenext
								  if rsS22.eof then exit for 
								  next
								  end if
								  rsS22.close
								  set rsS22=nothing
								  %>
                                   </div></td> 
								   
								  <td height="22"><div align="center">
									<%
								 set rsS32 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='2' and arr_course.theweek='3' order by arr_course.theweek"
								 rsS32.open sqlS,conn,1,1
								
								 if not (rsS32.eof and rsS32.bof) then
								 for i=0 to rsS32.recordcount
								 %>
								<%=rsS32("course")%>
                                  <%
								  rsS32.movenext
								  if rsS32.eof then exit for 
								  next
								  end if
								  rsS32.close
								  set rsS32=nothing
								  %>
                                   </div></td>  
								   
								   <td height="22"><div align="center">
									<%
								 set rsS41 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='2' and arr_course.theweek='4' order by arr_course.theweek"
								 rsS41.open sqlS,conn,1,1
								
								 if not (rsS41.eof and rsS41.bof) then
								 for i=0 to rsS41.recordcount
								 %>
								<%=rsS41("course")%>
                                  <%
								  rsS41.movenext
								  if rsS41.eof then exit for 
								  next
								  end if
								  rsS41.close
								  set rsS41=nothing
								  %>
                                   </div></td>   
								    <td height="22"><div align="center">
									<%
								 set rsS51 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='2' and arr_course.theweek='5' order by arr_course.theweek"
								 rsS51.open sqlS,conn,1,1
								
								 if not (rsS51.eof and rsS51.bof) then
								 for i=0 to rsS51.recordcount
								 %>
								<%=rsS51("course")%>
                                  <%
								  rsS51.movenext
								  if rsS51.eof then exit for 
								  next
								  end if
								  rsS51.close
								  set rsS51=nothing
								  %>
                                   </div></td>  
								  </tr>
								   <tr class="style10">
								   <td height="22"><div align="center">5.6��</div></td>
								   
                                    <td height="22"><div align="center">
									<%
								 set rsS11 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='3' and arr_course.theweek='1' order by arr_course.theweek"
								 rsS11.open sqlS,conn,1,1
								
								 if not (rsS11.eof and rsS11.bof) then
								 for i=0 to rsS11.recordcount
								 %>
								<%=rsS11("course")%>
                                  <%
								  rsS11.movenext
								  if rsS11.eof then exit for 
								  next
								  end if
								  rsS11.close
								  set rsS11=nothing
								  %>
                                   </div></td> 
								  
							
							                                    <td height="22"><div align="center">
									<%
								 set rsS21 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='3' and arr_course.theweek='2' order by arr_course.theweek"
								 rsS21.open sqlS,conn,1,1
								
								 if not (rsS21.eof and rsS21.bof) then
								 for i=0 to rsS21.recordcount
								 %>
								<%=rsS21("course")%>
                                  <%
								  rsS21.movenext
								  if rsS21.eof then exit for 
								  next
								  end if
								  rsS21.close
								  set rsS21=nothing
								  %>
                                   </div></td> 
								   
								  <td height="22"><div align="center">
									<%
								 set rsS31 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='3' and arr_course.theweek='3' order by arr_course.theweek"
								 rsS31.open sqlS,conn,1,1
								
								 if not (rsS31.eof and rsS31.bof) then
								 for i=0 to rsS31.recordcount
								 %>
								<%=rsS31("course")%>
                                  <%
								  rsS31.movenext
								  if rsS31.eof then exit for 
								  next
								  end if
								  rsS31.close
								  set rsS31=nothing
								  %>
                                   </div></td>  
								   
								   <td height="22"><div align="center">
									<%
								 set rsS41 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='3' and arr_course.theweek='4' order by arr_course.theweek"
								 rsS41.open sqlS,conn,1,1
								
								 if not (rsS41.eof and rsS41.bof) then
								 for i=0 to rsS41.recordcount
								 %>
								<%=rsS41("course")%>
                                  <%
								  rsS41.movenext
								  if rsS41.eof then exit for 
								  next
								  end if
								  rsS41.close
								  set rsS41=nothing
								  %>
                                   </div></td>   
								    <td height="22"><div align="center">
									<%
								 set rsS51 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='3' and arr_course.theweek='5' order by arr_course.theweek"
								 rsS51.open sqlS,conn,1,1
								
								 if not (rsS51.eof and rsS51.bof) then
								 for i=0 to rsS51.recordcount
								 %>
								<%=rsS51("course")%>
                                  <%
								  rsS51.movenext
								  if rsS51.eof then exit for 
								  next
								  end if
								  rsS51.close
								  set rsS51=nothing
								  %>
                                   </div></td>  
								  </tr>
								   <tr class="style10">
								   <td height="22"><div align="center">7.8��</div></td>
								   
                                    <td height="22"><div align="center">
									<%
								 set rsS11 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='4' and arr_course.theweek='1' order by arr_course.theweek"
								 rsS11.open sqlS,conn,1,1
								
								 if not (rsS11.eof and rsS11.bof) then
								 for i=0 to rsS11.recordcount
								 %>
								<%=rsS11("course")%>
                                  <%
								  rsS11.movenext
								  if rsS11.eof then exit for 
								  next
								  end if
								  rsS11.close
								  set rsS11=nothing
								  %>
                                   </div></td> 
								  
							
							                                    <td height="22"><div align="center">
									<%
								 set rsS21 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='4' and arr_course.theweek='2' order by arr_course.theweek"
								 rsS21.open sqlS,conn,1,1
								
								 if not (rsS21.eof and rsS21.bof) then
								 for i=0 to rsS21.recordcount
								 %>
								<%=rsS21("course")%>
                                  <%
								  rsS21.movenext
								  if rsS21.eof then exit for 
								  next
								  end if
								  rsS21.close
								  set rsS21=nothing
								  %>
                                   </div></td> 
								   
								  <td height="22"><div align="center">
									<%
								 set rsS31 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='4' and arr_course.theweek='3' order by arr_course.theweek"
								 rsS31.open sqlS,conn,1,1
								
								 if not (rsS31.eof and rsS31.bof) then
								 for i=0 to rsS31.recordcount
								 %>
								<%=rsS31("course")%>
                                  <%
								  rsS31.movenext
								  if rsS31.eof then exit for 
								  next
								  end if
								  rsS31.close
								  set rsS31=nothing
								  %>
                                   </div></td>  
								   
								   <td height="22"><div align="center">
									<%
								 set rsS41 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='4' and arr_course.theweek='4' order by arr_course.theweek"
								 rsS41.open sqlS,conn,1,1
								
								 if not (rsS41.eof and rsS41.bof) then
								 for i=0 to rsS41.recordcount
								 %>
								<%=rsS41("course")%>
                                  <%
								  rsS41.movenext
								  if rsS41.eof then exit for 
								  next
								  end if
								  rsS41.close
								  set rsS41=nothing
								  %>
                                   </div></td>   
								    <td height="22"><div align="center">
									<%
								 set rsS51 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&admin_academy&"' and arr_course.coursetime='4' and arr_course.theweek='5' order by arr_course.theweek"
								 rsS51.open sqlS,conn,1,1
								
								 if not (rsS51.eof and rsS51.bof) then
								 for i=0 to rsS51.recordcount
								 %>
								<%=rsS51("course")%>
                                  <%
								  rsS51.movenext
								  if rsS51.eof then exit for 
								  next
								  end if
								  rsS51.close
								  set rsS51=nothing
								  %>
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
                        <td width="60%" class="style3"><%=admin_account%>��</td>
                        <td>&nbsp;</td>
                      </tr>
					  <tr>
                        <td ><div align="center"> </div>
                            <div align="center"></div></td>
                        <td height="30" class="style2">����:<span class="style10"><%=rs2("admin_yx")%></span></td>
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
