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
dim job_ID
dim job_title,job_content,job_info
job_ID=session("job_ID")
job_title=trim(request("job_title"))
job_content=trim(request("job_content"))
job_info=trim(request("job_info"))
set rst=server.createobject("adodb.recordset")
sql="select * from job where job_id="&job_ID
rst.open sql,conn,1,3
%>
<%
    rst("job_title")=job_title
	rst("job_content")=job_content
	rst("job_info")=job_info
	rst.update
	rst.close
	set rst=nothing
	response.redirect "job_set.asp"
%>