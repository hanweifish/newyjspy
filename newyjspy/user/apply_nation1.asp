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
dim user_account,user_name,user_number,user_major,email,phone,user_tutor,user_address,user_code,sorts,score,school,major,pclb,applydate,user_Pyxz,user_Csrq
user_account=session("user_account")
sorts=trim(request("sorts"))
score=trim(request("score"))
email=trim(request("email"))
phone=trim(request("phone"))
school=trim(request("school"))
major=trim(request("major"))
pclb=trim(request("pclb"))
applydate=trim(request("applydate"))

set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,1
user_name=rs("user_name")
user_number=rs("user_number")
user_major=rs("user_major")
user_Pyxz=rs("user_Pyxz")
user_Csrq=rs("user_Csrq")
user_tutor=rs("user_tutor")
%>
<%
if user_Pyxz="定向"or user_Pyxz="委培" then 
Response.Write "<script> alert('您不符合申报条件!!');parent.window.history.go(-1);</script>"
Response.end
else
%>
<%
rs.close
set rs=nothing
%>
<%
set rs1=server.createobject("adodb.recordset")
sql="select * from apply_nation where user_number='"&user_number&"'"
rs1.open sql,conn,1,3
if not rs1.eof then
Response.Write "<script> alert('已经提交完毕，不需要重复提交！！');parent.window.history.go(-1);</script>"
Response.end
else
    rs1.addnew
    rs1("user_name")=user_name
    rs1("user_number")=user_number	
	rs1("user_major")=user_major
	rs1("user_tutor")=user_tutor
	rs1("user_Pyxz")=user_Pyxz
	rs1("user_Csrq")=user_Csrq	
	rs1("sorts")=sorts
	rs1("score")=score
	rs1("school")=school
	rs1("major")=major	
	rs1("email")=email
	rs1("pclb")=pclb
	rs1("phone")=phone
	rs1("applydate")=applydate	
	session("user_account")=user_account
	rs1.update
	rs1.close
	set rs1=nothing
	response.redirect "apply.asp"
end if
%>
<%
end if
%>
