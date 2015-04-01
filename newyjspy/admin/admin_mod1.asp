<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>
<%
if session("admin_account")="" then
Response.write"对不起，您还没有登陆或者无此权限！"
Response.end
end if
%>
<%
dim admin_number,admin_account,admin_pwd,admin_info,user_group,admin_ID
admin_ID=session("admin_ID")
admin_account=trim(request("admin_account"))
admin_pwd=trim(request("admin_pwd"))
admin_info=trim(request("admin_info"))
user_group=trim(request("user_group"))
set rs=server.createobject("adodb.recordset")
sql="select * from admin_info where admin_ID="&admin_ID
rs.open sql,conn,1,3
%>
<%
    rs("admin_account")=admin_account
	rs("admin_pwd")=admin_pwd
	rs("user_group")=user_group
	rs("admin_info")=admin_info
	rs.update
	rs.close
	set rs=nothing
	response.redirect "admin_info.asp"
%>