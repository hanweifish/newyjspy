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
dim course_id,theweek,doubleweek,coursetime
course_id=trim(request("course_id"))
theweek=trim(request("theweek"))
doubleweek=trim(request("doubleweek"))
coursetime=trim(request("coursetime"))

set rs=server.createobject("adodb.recordset")
sql="select * from arr_course where theweek="&theweek& "and coursetime="&coursetime
rs.open sql,conn,1,3
%>
<%
if not rs.eof then
if rs("doubleweek")="no" then
Response.Write "<script> alert('该课程已经存在！！');parent.window.history.go(-1);</script>"
Response.end
end if
else
    rs.addnew
    rs("course_id")=course_id
	rs("coursetime")=coursetime
	rs("theweek")=theweek
	rs("doubleweek")=doubleweek
	
	rs.update
	rs.close
	set rs=nothing
	response.redirect "arr_course.asp"
end if
%>