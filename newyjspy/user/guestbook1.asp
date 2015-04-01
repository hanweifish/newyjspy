<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("user_account")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>
<%
dim user_account,user_ID,title,content
title=trim(request("title"))
content=trim(request("content"))
user_account=trim(request("user_account"))
set rsu=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&session("user_account")&"'"
rsu.open sql,conn,1,1
set rs=server.createobject("adodb.recordset")
sql="select * from guestbook"
rs.open sql,conn,1,3
user_ID=rsu("user_ID")
rs.addnew
rs("user_time")=date()
rs("user_ID")=user_ID
rs("title")=title
rs("content")=content
rs.update
rs.close
set rs=nothing
response.write"<script>alert('发表留言成功！');document.location.href='user_index.asp';</script>"
%>
