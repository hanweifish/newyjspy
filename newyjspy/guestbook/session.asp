<%
if request.cookies("status")="" then
    Response.write"����û�е�½����½�󷽿����ԣ�"
	Response.end
end if
%>

<%
if session("user_account")="" then
    Response.write"����û�е�½����½�󷽿����ԣ�"
	Response.end
end if
%>
