<!--#include file="conn.asp"-->
<%
dim admin_account,admin_pwd
admin_account=trim(request("admin_account"))
admin_pwd=trim(request("admin_pwd"))
set rs=server.createobject("adodb.recordset")
sql="select * from teacher_info where admin_account='"&admin_account&"'" 
rs.open sql,conn,1,1

%>
<%
if not rs.eof then
	if rs("admin_pwd")<>admin_pwd then
		response.write "<script>alert('�Բ������벻��ȷ������������');document.location.href='../teacherlogin.asp';</script>"
		response.end
	else
		session("admin_account")=admin_account
		response.cookies("status")="statuson"
		response.redirect "admin_info.asp"
	end if
else
	response.write "<script>alert('�Բ��������߱�����ԱȨ�ޣ�');document.location.href='../teacherlogin.asp';</script>"
	response.end
end if
%>