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
user_number=trim(request("user_number"))
session("user_number")=user_number
train_academy=trim(request("train_academy"))
train_to=trim(request("train_to"))
train_info=trim(request("train_info"))

set rs2=server.createobject("adodb.recordset")
sql="select * from user_info where user_number='"&user_number&"'"
rs2.open sql,conn,1,1

if rs2("user_pyxz")="定向" or rs2("user_pyxz")="委培" then 
Response.Write "<script> alert('该同学不符合申报条件！');parent.window.history.go(-1);</script>"
else
if rs2.eof or rs2.bof then
Response.Write "<script> alert('该同学不存在！！');parent.window.history.go(-1);</script>"
else

dim user_ID
user_ID = rs2("user_ID")
set rs=server.createobject("adodb.recordset")
sql="select * from train_info where user_id="&user_ID
rs.open sql,conn,3,3
%>
<%
	if not(rs.eof and rs.bof) then
	Response.Write "<script>alert('此记录已经存在!')</script>" 
	train1=rs("train_ID")
    Response.Redirect "train_infomod.asp?train_ID="&train1
	else
    rs.addnew
	rs("user_ID")=rs2("user_ID")
	rs("train_to")=train_to
	rs("train_academy")=train_academy
	rs("train_info")=train_info
	
	rs.update
	rs.close
	set rs=nothing
    response.redirect "train_infoadd.asp"
	end if
end if
end if
	rs2.close
	set rs2=nothing
%>