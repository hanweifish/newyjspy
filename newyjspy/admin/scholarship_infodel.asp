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
dim course_id,page
page=trim(request("page"))
scholarship_academy=session("scholarship_academy")
scholarship_id=trim(request("scholarship_ID"))
set rs=server.createobject("adodb.recordset")
sql="select * from scholarship_info where scholarship_ID="&scholarship_id
rs.open sql,conn,1,3
%>

<%
rs.delete
rs.close
set rs=nothing
response.redirect "scholarship_set.asp"
%>