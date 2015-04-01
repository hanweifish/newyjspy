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
dim user_account,user_name,user_number,sqphone,zzsbldsq,user_Pyxz,zzsbld1,zzsbld2
user_account=session("user_account")
zzsbld1=trim(request("zzsbld1"))
sqphone=trim(request("sqphone"))
zzsbldsq=trim(request("zzsbldsq"))
zzsbld2=trim(request("zzsbld2"))
zzsbld3=trim(request("zzsbld3"))
zzsbld4=trim(request("zzsbld4"))

set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,3
if rs("user_zzsbld")="y" then
Response.Write "<script> alert('已经提交完毕，不需要重复提交！！');parent.window.history.go(-1);</script>"
Response.end
else
	rs("user_sqphone")=sqphone
	rs("user_zzsbld")="y"
	rs("user_zzsbld1")=zzsbld1
	rs("user_zzsbld2")=zzsbld2	
	rs("user_zzsbld3")=zzsbld3	
	rs("user_zzsbld4")=zzsbld4			
	rs("user_zzsbldsq")=zzsbldsq
end if

session("user_account")=user_account
rs.update
rs.close
set rs=nothing
response.redirect "apply.asp"

%>
