<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起ˇ您还没有登陆ˇ无此权ˇˇ"
	Response.end
end if
%>
<%
if session("admin_account")="" then
    Response.write"对不起ˇ您还没有登陆ˇ无此权ˇˇ"
	Response.end
end if
%>
<%
dim user_number,train_to,train_info,train_academy

train_id=trim(request("train_id"))
train_to=trim(request("train_to"))
train_info=trim(request("train_info"))

set rs=server.createobject("adodb.recordset")
sql="select * from train_info where train_id="&train_ID
rs.open sql,conn,1,3
%>
<%

	rs("train_to")=train_to
	rs("train_info")=train_info
	
	rs.update
	rs.close
	set rs=nothing
    response.redirect "train_info.asp"
	rsu.close
	set rsu=nothing
%>