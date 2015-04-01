<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>
<%
if session("admin_account")=""  then
Response.write"对不起，您还没有登陆或者无此权限！"
Response.end
end if
%>
<%
dim admin_number,admin_account,admin_pwd,admin_info,admin_ID,admin_academy,admin_yx
admin_ID=session("admin_ID")
admin_account=trim(request("admin_account"))
admin_pwd=trim(request("admin_pwd"))
admin_info=trim(request("admin_info"))
admin_yx=trim(request("admin_yx"))
admin_academy=trim(request("admin_academy"))
set rs=server.createobject("adodb.recordset")
sql="select * from teacher_info where admin_ID="&admin_ID
rs.open sql,conn,1,3
%>
<%
    rs("admin_account")=admin_account
	rs("admin_pwd")=admin_pwd
	rs("admin_yx")=admin_yx
	rs("admin_academy")=admin_academy
	rs("admin_info")=admin_info
	rs.update
	rs.close
	set rs=nothing
	response.redirect "teacher_info.asp"
%>