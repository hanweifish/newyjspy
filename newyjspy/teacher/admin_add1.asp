<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>
<%
if session("admin_account")="" or session("user_group")="subadmin" then
Response.write"�Բ�������û�е�½,���߲��߱�Ȩ�ޣ�"
Response.end
end if
%>
<%
dim admin_number,admin_account,admin_pwd,admin_info,user_group,admin_academy
admin_account=trim(request("admin_account"))
admin_pwd=trim(request("admin_pwd"))
admin_academy=trim(request("admin_academy"))
admin_info=trim(request("admin_info"))
user_group=trim(request("user_group"))
set rs=server.createobject("adodb.recordset")
sql="select * from admin_info where admin_account='"&admin_account&"'"
rs.open sql,conn,1,3
%>
<%
if not rs.eof then
Response.Write "<script> alert('���û��Ѿ����ڣ���');parent.window.history.go(-1);</script>"
Response.end
else
    rs.addnew
    rs("admin_account")=admin_account
	rs("admin_pwd")=admin_pwd
	rs("user_group")=user_group
	rs("admin_academy")=admin_academy
	rs("admin_info")=admin_info
	rs.update
	rs.close
	set rs=nothing
	response.redirect "admin_info.asp"
end if
%>