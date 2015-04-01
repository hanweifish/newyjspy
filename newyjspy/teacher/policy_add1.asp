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
dim today
today=Date
today=Year(today) & "-" & Right("0" & Month(today),2) & "-" & Right("0" & Day(today),2)
dim policy_title,policy_content,policy_info,policy_author
policy_title=trim(request("policy_title"))
policy_content=trim(request("policy_content"))
policy_author=trim(request("policy_author"))
policy_info=trim(request("policy_info"))
set rs=server.createobject("adodb.recordset")
sql="select * from policy "
rs.open sql,conn,1,3
%>
<%
    rs.addnew
    rs("policy_title")=policy_title
	rs("policy_content")=policy_content
	rs("policy_author")=policy_author
	rs("policy_time")=today
	rs.update
	rs.close
	set rs=nothing
	response.redirect "policy_set.asp"
%>