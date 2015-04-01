<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("admin_account")="" or session("user_group")="" then
Response.write"对不起，您还没有登陆，无此权限！"
Response.end
end if
%>


<%
dim user_ID,title,content,author
user_ID = trim(request("user_ID"))
title=trim(request("title"))
content=trim(request("content"))
author=trim(request("author"))

set rs=server.createobject("adodb.recordset")
sql="select * from reply"
rs.open sql,conn,1,3

rs.addnew
rs("user_ID")=user_ID
rs("title")=title
rs("content")=content
rs("author")=author
rs.update
rs.close
set rs=nothing
response.write"<script>alert('回复留言成功！');document.location.href='guestbook.asp';</script>"
%>
