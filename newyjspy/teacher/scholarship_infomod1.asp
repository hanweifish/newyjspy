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
dim user_number,scholarship_academy
user_number=trim(request("user_number"))
session("user_number")=user_number
scholarship_sorts=trim(request("scholarship_sorts"))
scholarship_prize=trim(request("scholarship_prize"))
scholarship_info=trim(request("scholarship_info"))
scholarship_academy=trim(request("admin_academy"))

set rs2=server.createobject("adodb.recordset")
sql="select * from user_info where user_number='"&user_number&"'"
rs2.open sql,conn,1,1

if rs2.eof or rs2.bof then
Response.Write "<script> alert('该同学不存在！！');parent.window.history.go(-1);</script>"
else

dim user_ID
user_ID = rs2("user_ID")
set rs=server.createobject("adodb.recordset")
sql="select * from scholarship_info where user_id="&user_ID
rs.open sql,conn,3,3
%>
<%
	rs("user_ID")=rs2("user_ID")
	rs("scholarship_sorts")=scholarship_sorts
	rs("scholarship_prize")=scholarship_prize
	rs("scholarship_info")=scholarship_info
	rs("scholarship_academy")=scholarship_academy
	
	rs.update
	rs.close
	set rs=nothing
    response.redirect "scholarship_info.asp"

end if
	rs2.close
	set rs2=nothing
%>