<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>
<%
if session("admin_account")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>
<%
dim tutor_id
tutor_id=trim(request("tutor_ID"))
set rs=server.createobject("adodb.recordset")
sql="select * from tutor where tutor_ID="&tutor_id
rs.open sql,conn,1,3
%>

<%
rs.delete
rs.close
set rs=nothing
response.redirect "tutor_set.asp"
%>