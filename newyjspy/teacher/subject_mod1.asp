<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("admin_account")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
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
period=trim(request("period"))
people=trim(request("people"))
term=trim(request("term"))
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
	rs("teachway")=teachway
	rs("period")=period	
	rs("people")=people
	rs.update
	rs.close
	set rs=nothing
	response.redirect "subject_set.asp?page="&page
%>