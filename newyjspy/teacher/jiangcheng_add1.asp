<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("admin_account")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>
<%
dim today
today=Date 
today=Year(today) & "-" & Right("0" & Month(today),2) & "-" & Right("0" & Day(today),2)
dim user_number, record
user_number=trim(request("user_number"))
record=trim(request("record"))

set rs=server.createobject("adodb.recordset")
sql="select * from jiangcheng where user_number= '"&user_number&"' and record='"&record&"'"
rs.open sql,conn,1,3
%>
<%
	if rs.RecordCount <> 0 then
	response.write "<script>alert('��¼�Ѵ���!'); parent.window.history.go(-1);</script>"
	end if
    rs.addnew
	rs("user_number")=user_number
	rs("record")=record
	rs.update
	rs.close
	set rs=nothing
	response.redirect "jiangcheng.asp"
%>