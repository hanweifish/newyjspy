<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>
<%
if session("admin_account")="" or session("user_group")<>"admin" then
Response.write"对不起，无此权限！"
Response.end
end if
%>
<%
dim user_id,NoncePage
NoncePage=trim(request("NoncePage"))
user_id=trim(request("ID"))
set rs=server.createobject("adodb.recordset")
sql="select * from apply_nation where ID="&user_id
rs.open sql,conn,1,3
%>
<%
rs.delete
rs.close
set rs=nothing
response.redirect "adminapply_nation.asp?page="&NoncePage
%>