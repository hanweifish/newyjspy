<%@CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="../encode.asp"-->
<!--#include file="../admin/session.asp"-->
<%
	dim forum_ID
	forum_ID = trim(request("forum_ID"))
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>留言版</title>
<link href="../style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style8 {font-size: 13px}
-->
</style>
</head>

<body>
<div align="center">
  <!--#include file = "top2.asp"-->
  <table width="840" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="20" background="../images/leftbk.jpg">&nbsp;</td>
      <td colspan="2"><div align="center">
        <table width="85%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="25"><div align="center"></div></td>
          </tr>
          <tr>
            <td height="50"><div align="left"><img src="../images/rotate.gif" width="11" height="11"><span class="style8">&nbsp;留 言 板 &gt;&gt; </span><span class="style1"><a href="admin_index.asp" target="_parent"><font class="style2">查 看 留 言</font></a></span></div></td>
          </tr>
          <tr>
            <td height="1" bgcolor="#FFE6B0"></td>
          </tr>
<%
set rs1 = Server.Createobject("Adodb.Recordset")
sql = "select * from reforum where forum_ID = "&forum_ID&" order by reforum_time" 
rs1.open sql,conn,1,1
%>
          <%
set rs2 = Server.Createobject("Adodb.Recordset")
sql = "select * from forum inner join user_info on forum.user_ID = user_info.user_ID where forum_ID = "&forum_ID 
rs2.open sql,conn,1,1
%>
          <tr>
            <td height="25"> <div align="left">by <%=rs2("user_account")%>&nbsp;&nbsp; <%=rs2("forum_time")%>&nbsp;&nbsp; <%=rs2("forum_title")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;
			<a href=admin_del.asp?forum_ID=<%=rs2("forum_ID")%>>删除</a>
			</div></td>
          </tr>
          <tr>
            <td><div align="left"><br>
            &nbsp;&nbsp;&nbsp;&nbsp;<%=HTMLEncode(rs2("forum_content"))%></div></td>
          </tr>
          <tr>
            <td height="1" bgcolor="#FFE6B0"></td>
          </tr>
<%
	if not (rs1.eof and rs1.bof) then
	for i=1 to rs1.recordcount
%>          
<%
	set rs3=Server.createobject("adodb.recordset")
	sql3="select * from user_info where user_ID = "&rs1("user_ID")
	rs3.open sql3,conn,1,1
%>		  
		  <tr>
            <td height="25"><div align="left">by <%=rs3("user_account")%>&nbsp;&nbsp; <%=rs1("reforum_time")%>&nbsp;&nbsp; RE"<%=rs2("forum_title")%>"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;
              <a href= admin_del.asp?forum_ID=<%=rs2("forum_ID")%>><font class="style2">删除</font></a>
</div></td>
          </tr>
          <tr>
            <td><div align="left"><br>
&nbsp;&nbsp;&nbsp;&nbsp;<%=HTMLEncode(rs1("reforum_content"))%></div></td>
          </tr>
          <tr>
            <td height="1" bgcolor="#FFE6B0"></td>
          </tr>
<%
	rs3.close
	set rs3=nothing
	rs1.movenext
	if rs1.eof or rs1.bof then exit for
	next
	rs1.close
	set rs1=nothing
	end if
	rs2.close
	set rs2=nothing
%>		  
          <tr>
            <td height="30"><div align="center"><a href="admin_index.asp">返 回</a>
              </div></td>
          </tr>
        </table>
          </div></td>
      <td width="20" background="../images/rightbk.jpg">&nbsp;</td>
    </tr>
  </table>
  
  <!--#include file = "bottom1.asp"-->
</div>
</body>
</html>
