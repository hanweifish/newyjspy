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
dim user_number,course
user_number=session("user_number")
course=session("course")
session("course")=course
session("user_number")=user_number
session("property")=property
%>
<%
dim sheet_id,score,sheet_info,year,property,term,tutor
sheet_id=session("sheet_id")
score=trim(request("score"))
sheet_info=trim(request("sheet_info"))
term=trim(request("term"))
tutor=trim(request("tutor"))
session("term")=term
session("tutor")=tutor
year=trim(request("year"))
property=trim(request("property"))
session("property")=property
session("year")=year
set rs=server.createobject("adodb.recordset")
sql="select * from sheet where sheet_ID="&sheet_id
rs.open sql,conn,1,3
%>
<%
	rs("score")=score
	rs("sheet_info")=sheet_info
	rs("year")=year
	rs("property")=property
	rs("term")=term
	rs("tutor")=tutor
	rs.update
	rs.close
	set rs=nothing
response.redirect "sheet_add.asp"
%>