<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>
<%
if session("admin_account")="" or session("user_group")<>"admin" then
Response.write"�Բ����޴�Ȩ�ޣ�"
Response.end
end if
%>
<%
dim site_id
site_id=trim(request("site_id"))
set rs=server.createobject("adodb.recordset")
sql="select * from user_site where site_id="&site_id
rs.open sql,conn,1,3
%>

<%
rs.delete
rs.close
set rs=nothing
response.redirect "site_set.asp"
%>