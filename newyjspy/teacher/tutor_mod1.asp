<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>
<%
if session("admin_account")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>
<%
dim tutor_ID,tutot_name,tutor_major,tutor_post,tutor_mord,tutor_acad,tutor_ptj,tutor_dir,tutor_proj
tutor_ID=session("tutor_ID")
tutor_name=trim(request("tutor_name"))
tutor_major=trim(request("tutor_major"))
tutor_post=trim(request("tutor_post"))
tutor_mord=trim(request("tutor_mord"))
tutor_acad=trim(request("tutor_acad"))
tutor_ptj=trim(request("tutor_ptj"))
tutor_dir=trim(request("tutor_dir"))
tutor_proj=trim(request("tutor_proj"))
set rs=server.createobject("adodb.recordset")
sql="select * from tutor where tutor_ID="&tutor_ID
rs.open sql,conn,1,3
%>
<%
    rs("tutor_name")=tutor_name
    rs("tutor_major")=tutor_major
	rs("tutor_post")=tutor_post
	rs("tutor_mord")=tutor_mord
	rs("tutor_acad")=tutor_acad
	rs("tutor_ptj")=tutor_ptj
	rs("tutor_dir")=tutor_dir
	rs("tutor_proj")=tutor_proj
rs.update
	rs.close
	set rs=nothing
	response.redirect "tutor_set.asp"
%>