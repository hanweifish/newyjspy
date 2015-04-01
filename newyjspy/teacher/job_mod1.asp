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
dim job_ID
dim job_title,job_content,job_info,job_author
job_ID=session("job_ID")
job_title=trim(request("job_title"))
job_content=trim(request("job_content"))
job_info=trim(request("job_info"))
job_author=trim(request("job_author"))
set rs=server.createobject("adodb.recordset")
sql="select * from job where job_id="&job_ID
rs.open sql,conn,1,3
%>
<%
    rs("job_title")=job_title
	rs("job_content")=job_content
	rs("job_info")=job_info
	rs("job_time")=date
	rs("job_author")=job_author
	rs.update
	rs.close
	set rs=nothing
	response.redirect "job_set.asp"
%>