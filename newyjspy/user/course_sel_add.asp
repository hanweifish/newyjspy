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
dim today
today=Date 
today=Year(today) & "-" & Right("0" & Month(today),2) & "-" & Right("0" & Day(today),2)
%>
<%
dim course_ID,user_ID
course_ID=trim(request("course_ID"))
%>
<%
if course_ID="" then
response.write"<script> alert('所选课程不能为空！');parent.window.history.go(-1);</script>"
else
%>
<%
dim startdate
set rscs=server.createobject("adodb.recordset")
sql="select * from course_set"
rscs.open sql,conn,1,1
startdate=rscs("startdate")
%>

<%
set rsu=server.createobject("adodb.recordset")
sql="select user_ID from user_info where user_account='"&session("user_account")&"'"
rsu.open sql,conn,1,1
user_ID=rsu("user_ID")
set rs=server.createobject("adodb.recordset")
sql="select course_sel.user_ID,course_sel.selTime,course_sel.course_ID from course_sel inner join user_info on course_sel.user_ID=user_info.user_ID where course_sel.selTime >'"&startdate&"' and  user_info.user_account='"&session("user_account")&"' and course_sel.course_ID="&course_ID
rs.open sql,conn,1,3
%>
<%
if not rs.eof then
Response.Write "<script> alert('您已经选了该课程！');parent.window.history.go(-1);</script>"
Response.end
else
    rs.addnew
	rs("course_ID")=course_ID
	rs("user_ID")=user_ID
	rs("selTime")=today
	rs.update
	rs.close
	set rs=nothing
	rsu.close
	set rsu=nothing
	response.redirect"course_sel.asp"
end if
end if
%>