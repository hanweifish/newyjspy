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
dim user_account,applydate
user_account=session("user_account")
apartdate=trim(request("apartdate"))

%>


<%
set rs1=server.createobject("adodb.recordset")
sql="select * from apply where user_number='"&user_account&"'"
rs1.open sql,conn,1,3


    rs1("apartdate")=apartdate
  
	session("user_account")=user_account
	rs1.update
	rs1.close
	set rs1=nothing
	response.redirect "apply.asp"

%>

