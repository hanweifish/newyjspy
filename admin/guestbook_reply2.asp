<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("admin_account")="" or session("user_group")="" then
Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
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
response.write"<script>alert('�ظ����Գɹ���');document.location.href='guestbook.asp';</script>"
%>
