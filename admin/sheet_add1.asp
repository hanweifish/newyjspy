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
dim today
today=Date 
today=Year(today) & "-" & Right("0" & Month(today),2) & "-" & Right("0" & Day(today),2)

if left(time(),2) = "下午" then
today = today&" "&CStr(CInt(left(right(time(),8),2))+12)&right(time(),6)
else
today = today&" "&right(time(),8)
end if


dim user_number,course,score,property,sheet_info,yearCount,course_ID,term,tutor,credit
user_number=trim(request("user_number"))
course_ID=trim(request("course_ID"))
score=trim(request("score"))
property=trim(request("property"))
sheet_info=trim(request("sheet_info"))
yearCount=trim(request("year"))
term=trim(request("term"))
tutor=trim(request("tutor"))

set rsSubject=server.createobject("adodb.recordset")
sql_rsSubject="select * from subject where course_ID="&course_ID
rsSubject.open sql_rsSubject,conn,1,1

course = rsSubject("course")
credit=rsSubject("credit")
course_yx=rsSubject("course_yx")
course_academy=rsSubject("course_academy")
session("course")=course
session("tutor")=tutor
session("term")=term
session("year")=yearCount
session("user_number")=user_number
set rsu=server.createobject("adodb.recordset")
sql="select * from user_info where user_number='"&user_number&"'"
rsu.open sql,conn,1,1
if rsu.eof or rsu.bof then
Response.Write "<script> alert('该同学不存在！！');parent.window.history.go(-1);</script>"
else
dim user_ID
user_ID = rsu("user_ID")
set rs=server.createobject("adodb.recordset")
sql="select * from sheet where user_ID="&user_ID&" and course_ID="&course_ID
rs.open sql,conn,3,3
%>
<%
	if not(rs.eof and rs.bof) then
	Response.Write "<script>alert('此记录已经存在!');history.go(-1)</script>" 
	else
    rs.addnew
    rs("course_ID")=course_ID
	rs("course")=course
	rs("user_ID")=rsu("user_ID")
	rs("score")=score
	rs("sheet_yx")=course_yx
	rs("sheet_academy")=course_academy	
	rs("property")=property
	rs("sheet_info")=sheet_info
	rs("sheet_time")=today
	rs("year")=yearCount
    rs("tutor")=tutor
	rs("term")=term
	session("property")=property
	
	if ISNumeric(rs("score"))  then
		if score>=60  then
		rs("sheet_credit")=credit								
		else
		rs("sheet_credit")=0	
		end if
	else 
	rs("sheet_credit")=credit	
	end if
	

    response.redirect "sheet_addadmin.asp?academy="&course_academy&""
	rs.update
	rs.close
	set rs=nothing
	end if
end if
	rsu.close
	set rsu=nothing
%>