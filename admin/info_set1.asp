<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>
<%
if session("admin_account")="" or session("user_group")<>"admin" then
Response.write"对不起，无此权限！"
Response.end
end if
%>
<%
dim user_ID,NoncePage
dim user_account,user_name,user_pwd,user_number,user_major,user_grade,user_mail,user_roomphone,user_mobile,user_tutor,user_labphone,user_homephone,user_address,user_code,user_sex,user_Csrq,user_info,user_Sfzh,user_Ksdw,user_Ksbh,user_Bmh,user_Ksfs,user_Bydw,user_Dxwpdw,user_Mz,user_Hf,user_Gj

NoncePage=trim(request("NoncePage"))
user_ID=session("user_ID")
user_account=trim(request("user_account"))
user_pwd=trim(request("user_pwd"))
user_name=trim(request("user_name"))
user_number=trim(request("user_number"))
user_major=trim(request("user_major"))
user_grade=trim(request("user_grade"))
user_mail=trim(request("user_mail"))
user_mobile=trim(request("user_mobile"))
user_bbs=trim(request("user_bbs"))
user_roomphone=trim(request("user_roomphone"))
user_tutor=trim(request("user_tutor"))
user_homephone=trim(request("user_homephone"))
user_address=trim(request("user_address"))
user_code=trim(request("user_code"))
user_sex=trim(request("user_sex"))
user_Csrq=trim(request("user_Csrq"))

user_Sfzh=trim(request("user_Sfzh"))
user_Ksdw=trim(request("user_Ksdw"))
user_Ksbh=trim(request("user_Ksbh"))
user_Bmh=trim(request("user_Bmh"))
user_Ksfs=trim(request("user_Ksfs"))
user_Bydw=trim(request("user_Bydw"))
user_Dxwpdw=trim(request("user_Dxwpdw"))
user_Mz=trim(request("user_Mz"))
user_Hf=trim(request("user_Hf"))
user_Gj=trim(request("user_Gj"))

user_info=trim(request("user_info"))
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_id="&user_ID
rs.open sql,conn,1,3
%>
<%
    rs("user_account")=user_account
	rs("user_pwd")=user_pwd
	rs("user_number")=user_number
	rs("user_name")=user_name
	rs("user_major")=user_major
	rs("user_grade")=user_grade
	rs("user_mail")=user_mail
	rs("user_mobile")=user_mobile
	rs("user_bbs")=user_bbs
	rs("user_roomphone")=user_roomphone
	rs("user_tutor")=user_tutor
	rs("user_address")=user_address
	rs("user_code")=user_code
	rs("user_homephone")=user_homephone
	rs("user_sex")=user_sex
	rs("user_Csrq")=user_Csrq
	
	rs("user_Sfzh")=user_Sfzh
	rs("user_Ksdw")=user_Ksdw
	rs("user_Ksbh")=user_Ksbh
	rs("user_Bmh")=user_Bmh	
	rs("user_Ksfs")=user_Ksfs
	rs("user_Bydw")=user_Bydw
	rs("user_Dxwpdw")=user_Dxwpdw
	rs("user_Mz")=user_Mz
	rs("user_Hf")=user_Hf	
	rs("user_Gj")=user_Gj
	
	rs("user_info")=user_info
	rs.update
	rs.close
	set rs=nothing
	response.redirect "admin_index.asp?page="&NoncePage
%>