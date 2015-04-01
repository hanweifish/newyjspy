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