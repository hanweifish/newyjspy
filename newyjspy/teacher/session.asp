<%
if request.cookies("status")="" then
    Response.write"您还没有登陆，登陆后方可留言！"
	Response.end
end if
%>

<%
if session("admin_account")="" then
    Response.write"您还没有登陆，登陆后方可留言！"
	Response.end
end if
%>
