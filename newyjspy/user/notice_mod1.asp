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
dim notice_ID
dim notice_title,notice_content,notice_info,notice_time
notice_ID=session("notice_ID")
notice_title=trim(request("notice_title"))
notice_content=trim(request("notice_content"))
notice_info=trim(request("notice_info"))
notice_time=trim(request("notice_time"))
set rst=server.createobject("adodb.recordset")
sql="select * from notice where notice_id="&notice_ID
rst.open sql,conn,1,3
%>
<%
    rst("notice_title")=notice_title
	rst("notice_content")=notice_content
	rst("notice_info")=notice_info
	rst("notice_time")=notice_time
	rst.update
	rst.close
	set rst=nothing
	response.redirect "notice_set.asp"
%>