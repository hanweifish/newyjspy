<!--#include file="conn.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from apply_nation"
rs.open sql,conn,1,3
%>

<%
if Not(rs.bof and rs.eof) then
for i=1 to rs.RecordCount

user_number=rs("user_number")
set rs1=server.createobject("adodb.recordset")
sql="select * from user_info where user_number='"&user_number&"'"
rs1.open sql,conn,1,1
user_major=rs1("user_major")
user_tutor=rs1("user_tutor")
rs1.close
set rs1=nothing
rs("user_tutor")=user_tutor
rs("user_major")=user_major
%>
<%
rs.update
rs.movenext
next
end if
rs.close
set rs=nothing
%>