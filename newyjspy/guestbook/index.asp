<%@CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="session.asp"-->


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
  <!--#include file = "top1.asp"-->
  <table width="840" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="20" background="../images/leftbk.jpg">&nbsp;</td>
      <td colspan="2"><div align="center">
        <table width="85%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="25"><div align="center"></div></td>
          </tr>
          <tr>
            <td height="50"><div align="left"><img src="../images/rotate.gif" width="11" height="11"><span class="style8">&nbsp;留 言 板 &gt;&gt; </span><span class="style1"><a href="post.asp" target="_parent">我要留言</a></span></div></td>
          </tr>
          <%
set rs = Server.Createobject("Adodb.Recordset")
sql = "select forum.forum_time,forum.forum_title,forum.forum_content,forum.forum_ID,user_info.user_account,user_info.user_ID from forum inner join user_info on forum.user_ID=user_info.user_ID order by forum.forum_time desc" 
rs.open sql,conn,1,1
%>
          <tr>
            <td><div align="center">
                <%
	if Not (rs.eof and rs.bof) then
		NumRecord = rs.recordcount
		rs.pagesize = 30
		NumPage = rs.Pagecount
		if request("page")=empty then
		NouncePage = 1
		elseif Cint(request("page")) < 1 then
		NouncePage = 1
		else
		NouncePage = request("page")
		end if
		if Cint(request("page"))>Cint(NumPage) then
		NouncePage = Numpage
		end if
	else
		NumPage = 0
		NumRecord = 0
		NouncePage = 0
	end if	
	%>
                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="80" height="24" bgcolor="#CCCCCC"><div align="left">作 者</div></td>
                    <td width="140" height="24" bgcolor="#CCCCCC"><div align="left">日 期</div></td>
                    <td width="430" height="24" bgcolor="#CCCCCC"><div align="left">标 题</div></td>
                    <td width="30" height="24" bgcolor="#CCCCCC"><div align="left">回 复</div></td>
                  </tr>
                  <%
		if Not(rs.eof and rs.bof) then
			rs.move(Cint(NouncePage)-1)*30,1
			for i=0 to rs.pagesize
		%>
                  <tr>
		<%
			set rs1 = Server.CreateObject("adodb.recordset")
			sql1 = "select * from forum inner join reforum on forum.forum_ID = reforum.forum_ID where forum.forum_ID="&rs("forum_ID")
			rs1.open sql1,conn,1,1
		%>
                    <%
		if i Mod 2 = 1 then
		%>
                    <td width="80" height="24"><div align="left"><%=rs("user_account")%></div></td>
                    <td width="140" height="24"><div align="left"><%=rs("forum_time")%></div></td>
                    <td width="430" height="24"><div align="left"><img src="../images/dot4.gif" width="11" height="11">&nbsp;<a href="forum_detail.asp?forum_ID=<%=rs("forum_ID")%>"><%=rs("forum_title")%></a></div></td>
                    <td width="30" height="24"><div align="left"><%=rs1.recordcount%></div></td>
                    <%
		else 
		%>
                    <td width="80" height="24" bgcolor="#F0F0F0"><div align="left"><%=rs("user_account")%></div></td>
                    <td width="140" height="24" bgcolor="#F0F0F0"><div align="left"><%=rs("forum_time")%></div></td>
                    <td width="430" height="24" bgcolor="#F0F0F0"><div align="left"><img src="../images/dot4.gif" width="11" height="11">&nbsp;<a href="forum_detail.asp?forum_ID=<%=rs("forum_ID")%>"><%=rs("forum_title")%></a></div></td>
                    <td width="30" height="24" bgcolor="#F0F0F0"><div align="left"><%=rs1.recordcount%></div></td>
                    <%
		end if
		%>
		<%
			rs1.close
			set rs1=nothing
		%>
                  </tr>
                  <%
			rs.movenext
			if rs.eof then exit for
			next
		else
			Response.Write("<tr><td height = '24' colspan='4'><div align = 'center'><marquee scrolldelay=120 behavior=alternate>请在此发表留言！</marquee></div></td></tr>")
		end if
		rs.close
		set rs=nothing
		%>
                  <tr>
                    <td height="25" colspan="4">&nbsp;</td>
                  </tr>
                  <tr>
                    <td height="20" colspan="4" background="./images/sline.gif"><div align="right">
                        <input type="hidden" name="page" value="<%=NouncePage%>">
                        <%
		  if NouncePage > 1 then
		  Response.Write("|<a href = 'index.asp?page=1'>首页</a>||<a href = 'index.asp?page="&NouncePage-1&"'>上一页</a>|&nbsp;")
		  else
		  Response.Write("|首页||上一页|&nbsp;")
		  end if
		  if Cint(NouncePage)< Cint(NumPage) then
		  Response.Write("|<a href = 'index.asp?page="&NouncePage+1&"'>下一页</a>||<a href = 'index.asp?page="&NumPage&"'>尾页</a>|&nbsp")
		  else
		  Response.Write("|下一页||尾页|")
		  end if
		  %>
&nbsp;页次：<%=NouncePage%>/<%=NumPage%> &nbsp;共<%=NumRecord%>条记录&nbsp; </div></td>
                  </tr>
                </table>
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
