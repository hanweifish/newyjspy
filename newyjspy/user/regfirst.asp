<%
dim user_account
user_account=session("user_account")
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,1
if rs("user_name")="unreg" or rs("user_number")="待添加" or rs("user_mail")="待添加" or rs("user_roomphone")="待添加" then
response.write "<script> alert('请先注册自己的个人信息！');parent.window.history.go(-1);</script>" 
response.end
end if
%>
