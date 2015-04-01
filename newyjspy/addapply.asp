<!--#include file="conn.asp"-->

<%
set rs=server.createobject("adodb.recordset")
sql="select * from apply"
rs.open sql,conn,1,1

for i=1 to rs.recordcount
number=rs("user_number")
yqby1=rs("user_yqby")
yqbysq=rs("user_yqby1")
sqphone=rs("user_roomphone")

set rs1=server.createobject("adodb.recordset")
sql="select * from user_info where user_number='"&number&"'"
rs1.open sql,conn,1,3
rs1("user_yqby")="y"
rs1("user_yqby1")=yqby1
rs1("user_yqbysq")=yqbysq
rs1("user_sqphone")=sqphone
rs1.update
rs1.close
set rs1=nothing

rs.movenext
next

rs.close
set rs=nothing
%>

