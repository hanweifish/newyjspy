<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>
<%
if session("admin_account")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>
<%
dim yanyi_bg,yanyi_end,yaner_bg,yaner_end,yansan_bg,yansan_end
yanyi_bg=trim(request("yanyi_bg"))
yanyi_end=trim(request("yanyi_end"))
yaner_bg=trim(request("yaner_bg"))
yaner_end=trim(request("yaner_end"))
yansan_bg=trim(request("yansan_bg"))
yansan_end=trim(request("yansan_end"))
set rs=server.createobject("adodb.recordset")
sql="select * from ynumber"
rs.open sql,conn,1,3
%>
<%
    rs("yanyi_bg")=yanyi_bg
	rs("yanyi_end")=yanyi_end
	rs("yaner_bg")=yaner_bg
	rs("yaner_end")=yaner_end
	rs("yansan_bg")=yansan_bg
	rs("yansan_end")=yansan_end
	rs.update
	rs.close
	set rs=nothing
	response.redirect "ynumber_set.asp"
%>