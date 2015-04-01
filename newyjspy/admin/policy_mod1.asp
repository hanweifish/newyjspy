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
dim policy_ID
dim policy_title,policy_content,policy_info,policy_author
policy_ID=session("policy_ID")
policy_title=trim(request("policy_title"))
policy_content=trim(request("policy_content"))
policy_author=trim(request("policy_author"))
policy_info=trim(request("policy_info"))
set rs=server.createobject("adodb.recordset")
sql="select * from policy where policy_id="&policy_ID
rs.open sql,conn,1,3
%>
<%
    rs("policy_title")=policy_title
	rs("policy_content")=policy_content
	rs("policy_info")=policy_info
	rs("policy_author")=policy_author
	rs.update
	rs.close
	set rs=nothing
	response.redirect "policy_set.asp"
%>