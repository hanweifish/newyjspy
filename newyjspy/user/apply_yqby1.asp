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
dim user_account,user_name,user_number,sqphone,yqbysq,user_Pyxz,yqby1
user_account=session("user_account")
yqby1=trim(request("yqby1"))
sqphone=trim(request("sqphone"))
yqbysq=trim(request("yqbysq"))

set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,3
if rs("user_yqby")="y" then
Response.Write "<script> alert('�Ѿ��ύ��ϣ�����Ҫ�ظ��ύ����');parent.window.history.go(-1);</script>"
Response.end
else
	rs("user_sqphone")=sqphone
	rs("user_yqby")="y"
	rs("user_yqby1")=yqby1
	rs("user_yqbysq")=yqbysq
end if

session("user_account")=user_account
rs.update
rs.close
set rs=nothing
response.redirect "apply.asp"

%>
