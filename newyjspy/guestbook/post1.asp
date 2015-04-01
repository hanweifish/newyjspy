<!--#include file ="conn.asp"-->
<!--#include file="session.asp"-->

<%
	dim today
	today=Date 
	today=Year(today) & "-" & Right("0" & Month(today),2) & "-" & Right("0" & Day(today),2)
%>
<%
	if left(time(),2) = "下午" then
	today = today&" "&CStr(CInt(left(right(time(),8),2))+12)&right(time(),6)
	else
	today = today&" "&right(time(),8)
	end if
%>
<%
	dim forum_title,forum_content
	forum_title=trim(request("forum_title"))
	forum_content=trim(request("forum_content"))
%>

<%
	set rs = Server.createobject("adodb.recordset")
	sql = "select * from forum"
	rs.open sql,conn,3,3
	set rs1 = Server.createobject("adodb.recordset")
	sql1 = "select * from user_info where user_account='"&session("user_account")&"'"
	rs1.open sql1,conn,1,1
%>

<%
	rs.addnew
	rs("forum_title")=forum_title
	rs("forum_content")=forum_content
	rs("forum_time")=today
	rs("user_ID")=rs1("user_ID")
	rs1.close
	set rs1=nothing
	rs.update
	rs.close
	set rs=nothing
%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="refresh" content="3;URL=index.asp">
<title>发表留言</title>
<link href="../style.css" rel="stylesheet" type="text/css">
</head>

<body >
<div align="center">
<!--#include file = "top1.asp"-->
<table width="800"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="center">
      <p>&nbsp;</p>
      <p><span>发表留言成功！<br>
          本页将在3秒后返回<br>
      如果您的浏览器没有反应，请<a href=index.asp><b>点击此处返回</b></a></span>
&nbsp;</p>
    </div></td>
  </tr>
</table>
</div>
</body>
</html>
