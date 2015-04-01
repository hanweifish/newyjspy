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
dim course,tutor,credit,term,teachway,coursenumber
course=trim(request("course"))
coursenumber=trim(request("coursenumber"))
tutor=trim(request("tutor"))
credit=trim(request("credit"))
term=trim(request("term"))
admin_yx=trim(request("admin_yx"))
admin_academy=trim(request("admin_academy"))
teachway=trim(request("teachway"))
period=trim(request("period"))
people=trim(request("people"))
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
	rs("course_academy")=admin_academy
	rs("course_yx")=admin_yx	
	rs("teachway")=teachway
	rs("period")=period	
	rs("people")=people
	rs.update
	rs.close
	set rs=nothing
	response.redirect "subject_set.asp"
end if
%>