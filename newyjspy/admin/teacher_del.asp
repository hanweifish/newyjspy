<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>
<%
if session("admin_account")="" then
Response.write"�Բ�������û�е�½,���߲��߱�Ȩ�ޣ�"
Response.end
end if
%>
<%
dim admin_id
admin_id=trim(request("admin_ID"))
set rs=server.createobject("adodb.recordset")
sql="select * from teacher_info where admin_ID="&admin_id
rs.open sql,conn,1,3
%>

<%
rs.delete
rs.close
set rs=nothing
response.redirect "teacher_info.asp"
%>