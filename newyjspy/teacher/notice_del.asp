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
page=trim(request("page"))
notice_ID=trim(request("notice_ID"))
set rs=server.createobject("adodb.recordset")
sql="select * from notice where notice_ID="&notice_id
rs.open sql,conn,1,3
%>

<%
rs.delete
rs.close
set rs=nothing
response.redirect "notice_set.asp?page="&page
%>