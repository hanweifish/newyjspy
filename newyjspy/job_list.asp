<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="conn.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select top 15 * from job where job_type = 'invite' order by job_time desc"
rs.open sql,conn,1,1
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

<body style="background:transparent">
                                <div align="center">
                                  <table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#A8BAFF">
                                    <TBODY>
                                      <tr>
                                        <td colspan="3" height="40"><div align="center"><font class="style3"> ---- 招 聘 信 息 ----</font></div></td>
                                      </tr>
                                      <tr>
                                        <td height="20" colspan="2"><div align="center"></div></td>
                                        <td width="60" height="20"><div align="center">点击数</div></td>
                                      </tr>
                                      <%
if not (rs.eof and rs.bof) then
for i=1 to rs.recordcount
if i mod 2 = 0 then
%>
                                      <tr>
                                        <td width="25" height="20"><img src="indeximages/arrow.gif" width="14" height="12" align="absmiddle"></td>
                                        <td width="636" height="20"><div align="left"><font class="style3"><a href=job_detail.asp?job_ID=<%=rs("job_ID")%> target="_blank"><%=left(rs("job_title"),16)%>...<span class="style10"><%=rs("job_time")%></span></a></font></div></td>
                                        <td width="60" height="20"><div align="center"><%=rs("job_click")%></div></td>
                                      </tr>
<%
else
%>
                                      <tr>
                                        <td width="25" height="20" bgcolor="#FFFFFF"><img src="indeximages/arrow.gif" width="14" height="12" align="absmiddle"></td>
                                        <td height="20" bgcolor="#FFFFFF"><div align="left"><font class="style3"><a href=job_detail.asp?job_ID=<%=rs("job_ID")%> target="_blank"><%=left(rs("job_title"),16)%>...<span class="style10"><%=rs("job_time")%></span></a></font></div></td>
                                        <td width="60" height="20" bgcolor="#FFFFFF"><div align="center"><%=rs("job_click")%></div></td>
                                     </tr>
<%
end if
%>
<%								  
rs.movenext
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
