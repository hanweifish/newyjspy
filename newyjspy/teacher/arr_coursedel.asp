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
course_id=trim(request("course_ID"))
set rs=server.createobject("adodb.recordset")
sql="select * from arr_course where arr_courseid="&arr_courseid
rs.open sql,conn,1,3
%>

<%
rs.delete
rs.close
set rs=nothing
response.redirect "arr._course.asp?
%>