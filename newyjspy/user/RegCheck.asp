<!--#include file="conn.asp"-->

<%
dim year,month,day,user_type,user_account,user_name,user_pwd,user_number,user_department,user_mail,user_phone,user_address,user_code,user_sex,user_birth,user_experience,user_resume,user_advantage,user_join
year=trim(request("year"))
month=trim(request("month"))
day=trim(request("day"))
user_type=trim(request("user_type"))
user_account=trim(request("user_account"))
user_pwd=trim(request("user_pwd"))
user_name=trim(request("user_name"))
user_number=trim(request("user_number"))
user_department=trim(request("user_department"))
user_mail=trim(request("user_mail"))
user_phone=trim(request("user_phone"))
user_address=trim(request("user_address"))
user_code=trim(request("user_code"))
user_sex=trim(request("user_sex"))
user_birth=year&"-"&month&"-"&day
user_experience=trim(request("user_experience"))
user_resume=trim(request("user_resume"))
user_advantage=trim(request("user_advantage"))
user_join=trim(request("user_join"))
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,3
%>
<%
if not rs.eof then
		response.write "<script>alert('此用户名已经存在，请选用其他的用户名！');history.go(-1);</script>"
		response.end
else
	rs.addnew
	rs("user_type")=user_type
    rs("user_account")=user_account
	rs("user_pwd")=user_pwd
	rs("user_number")=user_number
	rs("user_name")=user_name
	rs("user_department")=user_department
	rs("user_mail")=user_mail
	rs("user_phone")=user_phone
	rs("user_address")=user_address
	rs("user_code")=user_code
	rs("user_sex")=user_sex
	rs("user_birth")=user_birth
	rs("user_experience")=user_experience
	rs("user_resume")=user_resume
	rs("user_advantage")=user_advantage
	rs("user_join")=user_join
	session("user_account")=user_account
	response.cookies("status")="statuson"
	rs.update
	rs.close
	set rs=nothing
	response.redirect "user_index.asp"
end if
%>