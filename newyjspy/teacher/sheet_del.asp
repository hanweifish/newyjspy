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
dim sheet_id
sheet_id=trim(request("sheet_ID"))
set rs=server.createobject("adodb.recordset")
sql="select * from sheet where sheet_ID="&sheet_id
rs.open sql,conn,1,3
%>

<%
rs.delete
rs.close
set rs=nothing
response.redirect "sheet_add.asp"
%>