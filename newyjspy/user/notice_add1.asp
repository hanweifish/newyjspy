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
today=Date 
today=Year(today) & "-" & Right("0" & Month(today),2) & "-" & Right("0" & Day(today),2)
notice_title=trim(request("notice_title"))
notice_content=trim(request("notice_content"))
notice_author=rs("user_account")
notice_info=trim(request("notice_info"))
set rst=server.createobject("adodb.recordset")
sql="select * from notice "
rst.open sql,conn,1,3
%>
<%
    rst.addnew
    rst("notice_title")=notice_title
	rst("notice_content")=notice_content
	rst("notice_author")=notice_author
	rst("notice_info")=notice_info
	rst("notice_time")=today
	rst("user_ID")=rs("user_ID")
	rst.update
	rst.close
	set rst=nothing
	response.redirect "notice_set.asp"
%>