<!--#include file ="conn.asp"-->
<!--#include file="session.asp"-->


<%
	dim forum_title,forum_content,forum_ID
	forum_ID=trim(request("forum_ID"))
	forum_title=trim(request("forum_title"))
	forum_content=trim(request("forum_content"))
%>

<%
	set rs = Server.createobject("adodb.recordset")
	sql = "select * from forum where forum_ID="&forum_ID
	rs.open sql,conn,3,3
%>

<%
	rs("forum_title")=forum_title
	rs("forum_content")=forum_content
	rs.update
	rs.close
	set rs=nothing
%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="refresh" content="3;URL=forum_detail.asp?forum_ID=<%=forum_ID%>">
<title>修改留言</title>
<link href="../style.css" rel="stylesheet" type="text/css">
</head>

<body >
<div align="center">
<!--#include file = "top1.asp"-->
<table width="800"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="center">
      <p>&nbsp;</p>
      <p><span>修改留言成功！<br>
          本页将在3秒后返回<br>
      如果您的浏览器没有反应，请<a href=forum_detail.asp?forum_ID=<%=forum_ID%>><b>点击此处返回</b></a></span>
&nbsp;</p>
    </div></td>
  </tr>
</table>
</div>
</body>
</html>
