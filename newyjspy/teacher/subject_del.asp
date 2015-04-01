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
sql="select * from subject where course_ID="&course_id
rs.open sql,conn,1,3
%>

<%
rs.delete
rs.close
set rs=nothing
response.redirect "subject_set.asp?page="&page
%>