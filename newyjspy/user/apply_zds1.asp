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
dim user_account,user_name,user_number,sqphone,zdssq,user_Pyxz,zds1,zds2
user_account=session("user_account")
zds1=trim(request("zds1"))
sqphone=trim(request("sqphone"))
zdssq=trim(request("zdssq"))
zds2=trim(request("zds2"))
zds3=trim(request("zds3"))

set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,3
if rs("user_zds")="y" then
Response.Write "<script> alert('已经提交完毕，不需要重复提交！！');parent.window.history.go(-1);</script>"
Response.end
else
	rs("user_sqphone")=sqphone
	rs("user_zds")="y"
	rs("user_zds1")=zds1
	rs("user_zds2")=zds2
	rs("user_zds3")=zds3				
	rs("user_zdssq")=zdssq
end if

session("user_account")=user_account
rs.update
rs.close
set rs=nothing
response.redirect "apply.asp"

%>
