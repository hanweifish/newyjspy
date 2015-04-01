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
dim user_account,user_name,user_number,sqphone,fxsq,user_Pyxz,fx1,fx2
user_account=session("user_account")
fx1=trim(request("fx1"))
sqphone=trim(request("sqphone"))
fxsq=trim(request("fxsq"))
fx2=trim(request("fx2"))

set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,3
if rs("user_fx")="y" then
Response.Write "<script> alert('已经提交完毕，不需要重复提交！！');parent.window.history.go(-1);</script>"
Response.end
else
	rs("user_sqphone")=sqphone
	rs("user_fx")="y"
	rs("user_fx1")=fx1
	rs("user_fx2")=fx2	
	rs("user_fxsq")=fxsq
end if

session("user_account")=user_account
rs.update
rs.close
set rs=nothing
response.redirect "apply.asp"

%>
