<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("user_account")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>
<%
dim notice_ID
notice_ID=trim(request("notice_ID"))
set rst=server.createobject("adodb.recordset")
sql="select * from notice where notice_ID="&notice_id
rst.open sql,conn,1,3
%>

<%
rst.delete
rst.close
set rst=nothing
response.redirect "notice_set.asp"
%>