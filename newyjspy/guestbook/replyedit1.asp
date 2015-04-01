<!--#include file ="conn.asp"-->
<!--#include file="session.asp"-->

<%
	dim reforum_ID,forum_ID,reforum_content
	reforum_ID=trim(request("reforum_ID"))
	reforum_content=trim(request("reforum_content"))
	set rs = Server.createobject("adodb.recordset")
	sql = "select * from reforum where reforum_ID="&reforum_ID
	rs.open sql,conn,3,3
%>

<%
	forum_ID=rs("forum_ID")
	rs("reforum_content")=reforum_content
	rs.update
	rs.close
	set rs=nothing
%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="refresh" content="3;URL=forum_detail.asp?forum_ID=<%=forum_ID%>">
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
      如果您的浏览器没有反应，请<a href=forum_detail.asp?forum_ID=<%=forum_ID%>><b>点击此处返回</b></a></span>
&nbsp;</p>
    </div></td>
  </tr>
</table>
</div>
</body>
</html>
