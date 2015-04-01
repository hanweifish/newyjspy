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
dim user_account,user_name,user_number,sqphone,zzysq,user_Pyxz,zzy1,zzy2,zzy3
user_account=session("user_account")
zzy1=trim(request("zzy1"))
sqphone=trim(request("sqphone"))
zzysq=trim(request("zzysq"))
zzy2=trim(request("zzy2"))
zzy3=trim(request("zzy3"))
zzy4=trim(request("zzy4"))

set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,3
if rs("user_zzy")="y" then
Response.Write "<script> alert('已经提交完毕，不需要重复提交！！');parent.window.history.go(-1);</script>"
Response.end
else
	rs("user_sqphone")=sqphone
	rs("user_zzy")="y"
	rs("user_zzy1")=zzy1
	rs("user_zzy2")=zzy2
	rs("user_zzy3")=zzy3
	rs("user_zzy4")=zzy4				
	rs("user_zzysq")=zzysq
end if

session("user_account")=user_account
rs.update
rs.close
set rs=nothing
response.redirect "apply.asp"

%>
