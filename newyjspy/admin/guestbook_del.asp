<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>
<%
if session("admin_account")="" or session("user_group")<>"admin" then
Response.write"�Բ����޴�Ȩ�ޣ�"
Response.end
end if
%>
<%
dim guestbook_id
guestbook_id=trim(request("guestbook_ID"))
set rs=server.createobject("adodb.recordset")
sql="select * from guestbook where guestbook_ID="&guestbook_id
rs.open sql,conn,1,3
%>

<%
rs.delete
rs.close
set rs=nothing
response.redirect "guestbook.asp"
%>