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
dim tutor_id
tutor_id=trim(request("tutor_ID"))
set rs=server.createobject("adodb.recordset")
sql="select * from tutor where tutor_ID="&tutor_id
rs.open sql,conn,1,3
%>

<%
rs.delete
rs.close
set rs=nothing
response.redirect "tutor_set.asp"
%>