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
dim course_ID,page
dim course,tutor,credit,term
page=trim(request("page"))
course_ID=session("course_ID")
course=trim(request("course"))
tutor=trim(request("tutor"))
credit=trim(request("credit"))
term=trim(request("term"))
course_academy=trim(request("course_academy"))
teachway=trim(request("teachway"))
set rs=server.createobject("adodb.recordset")
sql="select * from subject where course_id="&course_ID
rs.open sql,conn,1,3
%>
<%
    rs("course")=course
	rs("tutor")=tutor
	rs("credit")=credit
	rs("term")=term
	rs("course_academy")=course_academy
	rs("teachway")=teachway
	rs.update
	rs.close
	set rs=nothing
	response.redirect "subject_set.asp?page="&page
%>