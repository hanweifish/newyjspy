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
dim user_account,user_name,user_number,sqphone,xxsq,user_Pyxz,xx1
user_account=session("user_account")
xx1=trim(request("xx1"))
sqphone=trim(request("sqphone"))
xxsq=trim(request("xxsq"))

set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,3
if rs("user_xx")="y" then
Response.Write "<script> alert('�Ѿ��ύ��ϣ�����Ҫ�ظ��ύ����');parent.window.history.go(-1);</script>"
Response.end
else
	rs("user_sqphone")=sqphone
	rs("user_xx")="y"
	rs("user_xx1")=xx1
	rs("user_xxsq")=xxsq
end if

session("user_account")=user_account
rs.update
rs.close
set rs=nothing
response.redirect "apply.asp"

%>
