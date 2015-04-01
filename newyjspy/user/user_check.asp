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
		response.write "<script>alert('对不起，密码不正确，请重新输入');history.go(-1);</script>"
		response.end
	else
		if rs("user_name")="unreg" or rs("user_number")="待添加" or rs("user_mail")="待添加" or rs("user_roomphone")="待添加" then
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
	response.write "<script>alert('对不起，你的用户名不存在，请与管理员联系！');document.location.href='../index.asp';</script>"
	response.end
end if
		
%>