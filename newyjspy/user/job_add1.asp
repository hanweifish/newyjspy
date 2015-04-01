<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("user_account")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>
<!--#include file="regfirst.asp"--> 
<%
dim today
dim job_title,job_content,job_info,job_author,job_time
today=Date 
today=Year(today) & "-" & Right("0" & Month(today),2) & "-" & Right("0" & Day(today),2)
job_title=trim(request("job_title"))
job_content=trim(request("job_content"))
job_author=rs("user_account")
job_info=trim(request("job_info"))
job_time=today
set rst=server.createobject("adodb.recordset")
sql="select * from job "
rst.open sql,conn,1,3
%>
<%
    rst.addnew
    rst("job_title")=job_title
	rst("job_content")=job_content
	rst("job_author")=job_author
	rst("job_info")=job_info
	rst("job_time")=job_time
	rst("user_ID")=rs("user_ID")
	rst.update
	rst.close
	set rst=nothing
	response.redirect "job_set.asp"
%>