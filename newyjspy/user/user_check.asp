<!--#include file="conn.asp"-->
<%
dim user_account,user_pwd
user_account=trim(request("user_account"))
user_pwd=trim(request("user_pwd"))
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,1

%>
<%
if not rs.eof then
	if rs("user_pwd")<>user_pwd then
		response.write "<script>alert('�Բ������벻��ȷ������������');history.go(-1);</script>"
		response.end
	else
		if rs("user_name")="unreg" or rs("user_number")="�����" or rs("user_mail")="�����" or rs("user_roomphone")="�����" then
			session("user_account")=user_account
			session("user_number")=rs("user_number")
			response.cookies("status")="statuson"
			response.redirect "info_reg.asp"
			response.end
		else
			session("user_account")=user_account
			session("user_number")=rs("user_number")
			response.cookies("status")="statuson"
			response.redirect "user_index.asp"
			response.end
		end if
	end if
else
	response.write "<script>alert('�Բ�������û��������ڣ��������Ա��ϵ��');document.location.href='../index.asp';</script>"
	response.end
end if
		
%>