<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("user_account")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
dim site_admin,site_pwd
site_admin=trim(request("site_admin"))
site_pwd=trim(request("site_pwd"))
set rs=server.createobject("adodb.recordset")
sql="select * from user_site where site_admin='"&site_admin&"'"
rs.open sql,conn,1,1

%>
<%
if not rs.eof then
	if rs("site_pwd")<>site_pwd then
		response.write "<script>alert('�Բ������벻��ȷ������������');document.location.href='index.asp';</script>"
		response.end
	else
		session("site_admin")=site_admin
		response.cookies("status")="statuson"
		response.redirect "site_admin.asp"
	end if
else
	response.write "<script>alert('�Բ�������û��������ڣ��������Ա��ϵ��');document.location.href='index.asp';</script>"
	response.end
end if
		
%>