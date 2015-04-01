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
dim jiangcheng_ID
jiangcheng_ID=trim(request("jiangcheng_ID"))
set rs=server.createobject("adodb.recordset")
sql="select * from jiangcheng where jiangcheng_ID="&jiangcheng_ID
rs.open sql,conn,1,3
%>

<%
rs.delete
rs.close
set rs=nothing
response.redirect "jiangcheng.asp"
%>