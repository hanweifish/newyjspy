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
dim today
today=Date 
today=Year(today) & "-" & Right("0" & Month(today),2) & "-" & Right("0" & Day(today),2)
dim job_title,job_content,job_info,job_author,job_type
job_title=trim(request("job_title"))
job_content=trim(request("job_content"))
job_info=trim(request("job_info"))
job_author=trim(request("job_author"))
job_type=trim(request("job_type"))
set rs=server.createobject("adodb.recordset")
sql="select * from job "
rs.open sql,conn,1,3
%>
<%
    rs.addnew
    rs("job_title")=job_title
	rs("job_content")=job_content
	rs("job_info")=job_info
	rs("job_time")=today
	rs("job_type")=job_type
	rs("job_author")=job_author
	rs.update
	rs.close
	set rs=nothing
	response.redirect "job_set.asp"
%>