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
dim user_id,NoncePage
NoncePage=trim(request("NoncePage"))
user_id=trim(request("user_ID"))
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_ID="&user_id
rs.open sql,conn,1,3
%>

<%
rs.delete
rs.close
set rs=nothing
response.redirect "admin_index.asp?page="&NoncePage
%>