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
dim notice_ID,page
dim notice_title,notice_content,notice_info,notice_author,notice_authority,notice_time
page=trim(request("page"))
notice_ID=session("notice_ID")
notice_title=trim(request("notice_title"))
notice_content=trim(request("notice_content"))
notice_author=trim(request("notice_author"))
notice_authority=trim(request("notice_authority"))
notice_info=trim(request("notice_info"))
notice_time=trim(request("notice_time"))
set rs=server.createobject("adodb.recordset")
sql="select * from notice where notice_id="&notice_ID
rs.open sql,conn,1,3
%>
<%
    rs("notice_title")=notice_title
	rs("notice_content")=notice_content
	rs("notice_author")=notice_author
	rs("notice_authority")=notice_authority
	rs("notice_info")=notice_info
	rs("notice_time")=notice_time
	rs.update
	rs.close
	set rs=nothing
	response.redirect "notice_set.asp?page="&page
%>