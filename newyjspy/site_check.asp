<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("user_account")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
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
		response.write "<script>alert('对不起，密码不正确，请重新输入');document.location.href='index.asp';</script>"
		response.end
	else
		session("site_admin")=site_admin
		response.cookies("status")="statuson"
		response.redirect "site_admin.asp"
	end if
else
	response.write "<script>alert('对不起，你的用户名不存在，请与管理员联系！');document.location.href='index.asp';</script>"
	response.end
end if
		
%>