<%
dim user_account
user_account=session("user_account")
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,1
if rs("user_name")="unreg" or rs("user_number")="�����" or rs("user_mail")="�����" or rs("user_roomphone")="�����" then
response.write "<script> alert('����ע���Լ��ĸ�����Ϣ��');parent.window.history.go(-1);</script>" 
response.end
end if
%>
