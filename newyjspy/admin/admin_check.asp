<!--#include file="conn.asp"-->
<%
dim admin_account,admin_pwd
admin_account=trim(request("admin_account"))
admin_pwd=trim(request("admin_pwd"))
set rs=server.createobject("adodb.recordset")
sql="select * from admin_info where admin_account='"&admin_account&"'" 
rs.open sql,conn,1,1

%>
<%
if not rs.eof then
	if rs("admin_pwd")<>admin_pwd then
		response.write "<script>alert('�Բ������벻��ȷ������������');document.location.href='../adminlogin.asp';</script>"
		response.end
	else
		session("admin_account")=admin_account
		session("user_group")=rs("user_group")
		response.cookies("status")="statuson"
		if rs("user_group")<>"subadmin" then
		response.redirect "info_search.asp"
		else
		response.redirect "notice_set.asp"
		end if
	end if
else
	response.write "<script>alert('�Բ��������߱�����ԱȨ�ޣ�');document.location.href='../adminlogin.asp';</script>"
	response.end
end if
%>