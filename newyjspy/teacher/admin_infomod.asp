<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>
<%
if session("admin_account")="" then
    response.write "<script>alert('�Բ�������û�е�½���޴�Ȩ�ޣ�')</script>"
	Response.end
end if
%>
<%
dim admin_pwd,admin_pwd1,admin_account
admin_account=session("admin_account")
admin_pwd=trim(request("admin_pwd"))
admin_pwd1=trim(request("admin_pwd1"))
set rs=server.createobject("adodb.recordset")
sql="select * from teacher_info where admin_account='"&admin_account&"'"
rs.open sql,conn,1,3
%>
<%
if admin_pwd<>admin_pwd1 then
		response.write "<script>alert('������������벻ƥ�䣡');document.location.href='admin_info.asp';</script>"
		response.end
	else
	rs("admin_pwd")=admin_pwd
	rs.update
	rs.close
	set rs=nothing
	response.redirect "admin_info.asp"
end if
%>
