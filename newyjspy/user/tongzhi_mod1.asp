<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("user_account")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>
<!--#include file="regfirst.asp"--> 

<%
dim tongzhi_ID
dim tongzhi_title,tongzhi_content,tongzhi_info,tongzhi_time
tongzhi_ID=session("tongzhi_ID")
tongzhi_title=trim(request("tongzhi_title"))
tongzhi_content=trim(request("tongzhi_content"))
tongzhi_info=trim(request("tongzhi_info"))
tongzhi_time=trim(request("tongzhi_time"))
set rst=server.createobject("adodb.recordset")
sql="select * from tongzhi where tongzhi_id="&tongzhi_ID
rst.open sql,conn,1,3
%>
<%
    rst("tongzhi_title")=tongzhi_title
	rst("tongzhi_content")=tongzhi_content
	rst("tongzhi_info")=tongzhi_info
	rst("tongzhi_time")=tongzhi_time
	rst.update
	rst.close
	set rst=nothing
	response.redirect "tongzhi_set.asp"
%>