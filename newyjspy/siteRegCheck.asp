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
dim site_name,site_admin,site_pwd,site_pwd1,site_info,site_url
site_admin=trim(request("site_admin"))
site_pwd=trim(request("site_pwd"))
site_pwd1=trim(request("site_pwd1"))
site_name=trim(request("site_name"))
site_url=trim(request("site_url"))
site_info=trim(request("site_info"))

set rs=server.createobject("adodb.recordset")
sql="select * from user_site where site_admin='"&site_admin&"'"
rs.open sql,conn,1,3
%>
<%
if not rs.eof then
Response.Write "<script> alert('该用户主页已经存在！！');parent.window.history.go(-1);</script>"
Response.end
end if
if site_pwd<>site_pwd1 then
Response.Write"<script>alert('两次输入的密码不匹配！');parent.window.history.go(-1)</script>"
else
    rs.addnew
	rs("site_name")=site_name
	rs("site_admin")=site_admin
	rs("site_pwd")=site_pwd
	rs("site_url")=site_url
	rs("site_info")=site_info
	response.cookies("status")="statuson"
	session("site_admin")=site_admin
	rs.update
	rs.close
	set rs=nothing
	response.redirect "site_admin.asp"
end if
%>