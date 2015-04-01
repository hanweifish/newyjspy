<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("admin_account")=""then
Response.write"对不起，您还没有登陆，无此权限！"
Response.end
end if
%>

<%
dim admin_account
admin_account=session("admin_account")
dim admin_academy
academy=trim(request("academy"))
%>



<script language="javascript">
	function checkcourse(form)
	{
		if (document.form.user_number.value=="")
		{
			alert("请输入学生学号！");
		}
		else if (document.form.score.value=="")
		{
			alert("请输入成绩！");
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
                               
                                
                                <td><div align="center" class="style10">排课系统</div></td>
                               
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
                              <td><div align="center" class="style10"><%=academy%></div></td>
                            </tr>
                            <tr>
                              <td><div align="center">
                                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                
								  <tr class="style10">
                                     <td height="22" width=50><div align="center"></div></td>
									<td height="22"><div align="center">星期一</div></td>
                                    <td><div align="center">星期二</div></td>
                                    <td><div align="center">星期三</div></td>
									<td><div align="center">星期四</div></td>
                                    <td><div align="center">星期五</div></td>
                                   
								  </tr>
								 
								  <tr class="style10">
								   <td height="22"><div align="center">1.2节</div></td>
								   
                                    <td height="22"><div align="center">
									<%
								 set rsS11 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='1' and arr_course.theweek='1' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='1' and arr_course.theweek='2' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='1' and arr_course.theweek='3' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='1' and arr_course.theweek='4' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='1' and arr_course.theweek='5' order by arr_course.theweek"
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
								   <td height="22"><div align="center">3.4节</div></td>
								   
                                    <td height="22"><div align="center">
									<%
								 set rsS21 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='2' and arr_course.theweek='1' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='2' and arr_course.theweek='2' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='2' and arr_course.theweek='3' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='2' and arr_course.theweek='4' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='2' and arr_course.theweek='5' order by arr_course.theweek"
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
								   <td height="22"><div align="center">5.6节</div></td>
								   
                                    <td height="22"><div align="center">
									<%
								 set rsS11 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='3' and arr_course.theweek='1' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='3' and arr_course.theweek='2' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='3' and arr_course.theweek='3' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='3' and arr_course.theweek='4' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='3' and arr_course.theweek='5' order by arr_course.theweek"
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
								   <td height="22"><div align="center">7.8节</div></td>
								   
                                    <td height="22"><div align="center">
									<%
								 set rsS11 = Server.CreateObject("adodb.recordset")
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='4' and arr_course.theweek='1' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='4' and arr_course.theweek='2' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='4' and arr_course.theweek='3' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='4' and arr_course.theweek='4' order by arr_course.theweek"
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
								 sqlS="select   arr_course.course_id,arr_course.coursetime,arr_course.theweek,arr_course.doubleweek,subject.course_id,subject.course,subject.course_academy from subject inner join arr_course on arr_course.course_id=subject.course_id  where subject.course_academy='"&academy&"' and arr_course.coursetime='4' and arr_course.theweek='5' order by arr_course.theweek"
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
                        <td width="60%" class="style3"><%=admin_account%>：</td>
                        <td>&nbsp;</td>
                      </tr>
					  
                      <tr>
                        <td ><div align="center"> </div>
                            <div align="center"></div></td>
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
