<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("user_account")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>
<%
dim job_ID
job_ID=trim(request("job_ID"))
set rst=server.createobject("adodb.recordset")
sql="select * from job where job_ID="&job_id
rst.open sql,conn,1,3
%>

<%
rst.delete
rst.close
set rst=nothing
response.redirect "job_set.asp"
%>