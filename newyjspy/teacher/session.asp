<%
if request.cookies("status")="" then
    Response.write"����û�е�½����½�󷽿����ԣ�"
	Response.end
end if
%>

<%
if session("admin_account")="" then
    Response.write"����û�е�½����½�󷽿����ԣ�"
	Response.end
end if
%>
