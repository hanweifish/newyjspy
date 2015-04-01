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
dim user_account,user_name,user_number,sqphone,txsq,user_Pyxz,tx1
user_account=session("user_account")
tx1=trim(request("tx1"))
sqphone=trim(request("sqphone"))
txsq=trim(request("txsq"))

set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,3
if rs("user_tx")="y" then
Response.Write "<script> alert('已经提交完毕，不需要重复提交！！');parent.window.history.go(-1);</script>"
Response.end
else
	rs("user_sqphone")=sqphone
	rs("user_tx")="y"
	rs("user_tx1")=tx1
	rs("user_txsq")=txsq
end if

session("user_account")=user_account
rs.update
rs.close
set rs=nothing
response.redirect "apply.asp"

%>
