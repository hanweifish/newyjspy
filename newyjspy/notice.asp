<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="conn.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select top 15 * from notice order by notice_time desc"
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
                          <table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td valign="top"><div align="center">
                                <table width="98%" border="0" cellpadding="0" cellspacing="0" bordercolor="#A8BAFF">
<%
if not (rs.eof and rs.bof) then
for i=1 to rs.recordcount
%>
                                  <tr>
                                    <td width="20" height="25"><img src="indeximages/arrow.gif" width="13" height="11" align="absmiddle"></td>
                                    <td height="25"><div align="left"><font class="style2" style="cursor:hand" onclick="MM_openBrWindow('notice_detail.asp?notice_id=<%=rs("notice_ID")%>','通知','scrollbars=yes,resizable=yes,width=650')"><%=left(rs("notice_title"),25)%>...<span class="style10"><%=rs("notice_time")%>&nbsp;(点击<%=rs("notice_click")%>次)</span></font></div></td>
                                  </tr>
<%
rs.movenext
next
%>
                                </table>
                              </div></td>
                            </tr>
</table>
<%
else
%>
                          <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#A8BAFF">
                              <tr>
                                <td><div align="center"><font class="style3" >暂时没有新的通知发布！</font></div></td>
                              </tr>
                          </table>
<%
	end if
%>

</body>
</html>
