<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="conn.asp"-->
<%	
set rsjob=server.createobject("adodb.recordset")
jobsql="select * from job where job_type = 'policy' order by job_time desc"
rsjob.open jobsql,conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<style type="text/css">
<!--
.style10 {font-size: 10}
-->
</style>
</head>

<body bgcolor="">
                                <div align="center">
                                  <table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#A8BAFF">
                                    <TBODY>
                                      <tr>
                                        <td height="20" colspan="2"><div align="center"></div></td>
                                        <td width="65" height="20"><div align="center">点击数</div></td>
                                      </tr>
                                      <%
if not (rsjob.eof and rsjob.bof) then
for i=1 to rsjob.recordcount
if i mod 2 = 0 then
%>
                                      <tr>
                                        <td width="25" height="20"><img src="indeximages/triangle.gif" width="6" height="8" align="absmiddle"></td>
                                        <td width="631" height="20"><div align="left"><font class="style3"><a href=job_detail.asp?job_ID=<%=rsjob("job_ID")%> target="_blank"><%=left(rsjob("job_title"),16)%>...<span class="style10"><%=rsjob("job_time")%></span></a></font></div></td>
                                        <td width="65" height="20"><div align="center"><%=rsjob("job_click")%></div></td>
                                      </tr>
<%
else
%>
                                      <tr>
                                        <td width="25" height="20" bgcolor="#FFFFFF"><img src="indeximages/triangle.gif" width="6" height="8" align="absmiddle"></td>
                                        <td height="20" bgcolor="#FFFFFF"><div align="left"><font class="style3"><a href=job_detail.asp?job_ID=<%=rsjob("job_ID")%> target="_blank"><%=left(rsjob("job_title"),16)%>...<span class="style10"><%=rsjob("job_time")%></span></a></font></div></td>
                                        <td width="65" height="20" bgcolor="#FFFFFF"><div align="center"><%=rsjob("job_click")%></div></td>
                                      </tr>
<%
end if
%>
<%								  
rsjob.movenext
next
%>
<%
else
%>
                                      <tr>
                                        <td colspan="3"><div align="center"><font class="style3" >暂时没有新的招聘信息发布！</font></div></td>
                                      </tr>
<%
end if
%>
                                    </TBODY>
                                  </TABLE>
                                </div>
</body>
</html>
