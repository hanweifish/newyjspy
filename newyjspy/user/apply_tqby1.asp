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
dim user_account,user_name,user_number,sqphone,tqbysq,user_Pyxz,tqby1
user_account=session("user_account")
tqby1=trim(request("tqby1"))
sqphone=trim(request("sqphone"))
tqbysq=trim(request("tqbysq"))

set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,3
if rs("user_tqby")="y" then
Response.Write "<script> alert('�Ѿ��ύ��ϣ�����Ҫ�ظ��ύ����');parent.window.history.go(-1);</script>"
Response.end
else
	rs("user_sqphone")=sqphone
	rs("user_tqby")="y"
	rs("user_tqby1")=tqby1
	rs("user_tqbysq")=tqbysq
end if

session("user_account")=user_account
rs.update
rs.close
set rs=nothing
response.redirect "apply.asp"

%>
