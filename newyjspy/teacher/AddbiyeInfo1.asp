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
	dim user_number, page, jxsx, bylw, zdjs, dbsj, dbjg, byrq, wphm, xwhm, fzrqm, jwyqm
	user_number = trim(request("user_number"))
	page = trim(request("page"))
	jxsx = trim(request("jxsx"))
	bylw = trim(request("bylw"))
	zdjs = trim(request("zdjs"))
	dbsj = trim(request("dbsj"))
	dbjg = trim(request("dbjg"))
	byrq = trim(request("byrq"))
	wphm = trim(request("wphm"))
	xwhm = trim(request("xwhm"))
	fzrqm = trim(request("fzrqm"))
	jwyqm = trim(request("jwyqm"))
	set rsBiyeInfo = Server.CreateObject("Adodb.Recordset")
	sql_rsBiyeInfo = "Select * from biyeInfo where user_number = '"&user_number&"'"
	rsBiyeInfo.Open sql_rsBiyeInfo,conn,3,3
%>
<%
	if rsBiyeInfo.RecordCount = 0 then
	rsBiyeInfo.Addnew
	rsBiyeInfo("user_number") = user_number
	end if
    rsBiyeInfo("jxsx")=jxsx
	rsBiyeInfo("bylw")=bylw
	rsBiyeInfo("zdjs")=zdjs
	rsBiyeInfo("dbsj")=dbsj
	rsBiyeInfo("dbjg")=dbjg
	rsBiyeInfo("byrq")=byrq
	rsBiyeInfo("wphm")=wphm
	rsBiyeInfo("xwhm")=xwhm
	rsBiyeInfo("fzrqm")=fzrqm
	rsBiyeInfo("jwyqm")=jwyqm
	
	rsBiyeInfo.update
	rsBiyeInfo.close
	set rsBiyeInfo=nothing
	response.write "<script>alert('编辑成功!') </script>"
	response.redirect "Admin_index.asp?page="&page
%>