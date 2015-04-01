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
dim course,tutor,credit,term,teachway,course_academy
course=trim(request("course"))
tutor=trim(request("tutor"))
credit=trim(request("credit"))
term=trim(request("term"))
coursenumber=trim(request("coursenumber"))
course_academy=trim(request("course_academy"))
teachway=trim(request("teachway"))

set rs2=server.createobject("adodb.recordset")
sql="select * from teacher_info where admin_academy='"&course_academy&"'"
rs2.open sql,conn,1,1
course_yx=rs2("admin_yx")

set rs=server.createobject("adodb.recordset")
sql="select * from subject where tutor='"&tutor&"' and credit = '"&credit&"' and course='"&course&"'"
rs.open sql,conn,1,3
%>

<%
if not rs.eof then
Response.Write "<script> alert('该课程已经存在！！');parent.window.history.go(-1);</script>"
Response.end
else
    rs.addnew
    rs("course")=course
	rs("tutor")=tutor
	rs("credit")=credit
	rs("term")=term
	rs("teachway")=teachway
    rs("course_academy")=course_academy
	rs("course_yx")=course_yx
	rs.update
	rs.close
	set rs=nothing
    rs2.close
	set rs2=nothing
	response.redirect "subject_set.asp"
end if
%>