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
dim course_ID
dim course_name,course_number,course_tutor,course_credit,course_term,course_time,course_info
course_ID=session("course_ID")
course_name=trim(request("course_name"))
course_number=trim(request("course_number"))
course_tutor=trim(request("course_tutor"))
course_credit=trim(request("course_credit"))
course_term=trim(request("course_term"))
course_time=trim(request("course_time"))
course_info=trim(request("course_info"))
set rs=server.createobject("adodb.recordset")
sql="select * from course where course_id="&course_ID
rs.open sql,conn,1,3
%>
<%
    rs("course_name")=course_name
	rs("course_number")=course_number
	rs("course_tutor")=course_tutor
	rs("course_credit")=course_credit
	rs("course_term")=course_term
	rs("course_time")=course_time
	rs("course_info")=course_info
	rs.update
	rs.close
	set rs=nothing
	response.redirect "course_set.asp"
%>