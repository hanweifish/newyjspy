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
dim course_selid
course_selid=trim(request("course_selID"))
set rs=server.createobject("adodb.recordset")
sql="select * from course_sel where course_selID="&course_selid
rs.open sql,conn,1,3
%>
<%
rs.delete
rs.close
set rs=nothing
response.redirect"course_sel.asp"
%>