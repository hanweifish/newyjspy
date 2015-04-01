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
dim admin_account
admin_account=session("admin_account")
%>

<%
set rs=server.createobject("adodb.recordset")
sql="select * from teacher_info where admin_account='"&admin_account&"'"
rs.open sql,conn,1,1
admin_academy=rs("admin_academy")
%>
<head>
<script language="javascript">
window.status="欢迎访问南京大学物理系研究生管理信息系统！"
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>研究生信息管理系统</title>
<link href="../style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style10 {font-size: 12px;
	color: #004080;
}
-->
</style>
<style type="text/css">
<!--
.style12 {color: #006699;
	font-size: 12px;
}
-->
</style>
<style type="text/css">
<!--
.style14 {color: #FF6633;
	font-size: 12px;
}
.style15 {font-size: 11px}
-->
</style>
<style type="text/css">
<!--
.style16 {color: #FF6600}
-->
</style>
</head>
<body>
<table width="90%" height="155" cellpadding="0"  cellspacing="0">
                              <tr>
                                <td><div align="center">
                                    <table width="100%"  cellspacing="0" cellpadding="0">
                                      <tr>
                                        <td height="25"></td>
                                      </tr>
                                      <tr>
                                        <td><form action="train_infoadd.asp" method="post" name="form" onSubmit="return checkform(form)">
                                        </form></td>
                                      </tr>
                                    </table>
                                </div></td>
                              </tr>
                              <tr>
                                <td><div align="center">
                                    <table width="100%"  border="0" cellpadding="0" cellspacing="0">
                                      <tr>
                                        <td height="25" colspan="6"><table width="100%"  border="1" cellspacing="0" cellpadding="0">
                                          <%
								 set rsS = Server.CreateObject("adodb.recordset")
								 sqlS="select user_info.user_name,user_info.user_number,user_info.user_yx,user_info.user_pyxz,train_info.train_academy,train_info.train_info,train_info.train_ID,train_info.train_to from train_info inner join user_info on train_info.user_ID=user_info.user_ID where train_info.train_academy='"&admin_academy&"' order by train_info.train_id desc"
								 rsS.open sqlS,conn,1,1
								 
								  if Not(rsS.bof and rsS.eof) then
			NumRecord=rsS.recordcount
			rsS.pagesize=50
			NumPage=rsS.Pagecount
			if request("page")=empty then 
			NoncePage=1
		else
		if Cint(request("page"))<1 then
			NoncePage=1
		else
			NoncePage=Trim(request("page"))
		end if
		if Cint(Trim(request("page")))>Cint(NumPage) then NoncePage=NumPage
	end if
else
	NumRecord=0
	NumPage=0
	NoncePage=0
	end if
								 
								 %>
                                          <tr class="style10">
                                            <td height="22"><div align="center">姓名</div></td>
                                            <td><div align="center">学号</div></td>
                                            <td><div align="center">乘车目的地</div></td>
                                            <td><div align="center">院系</div></td>
											<td><div align="center">培养性质</div></td>
                                            <td><div align="center">附加信息</div></td>
                                          </tr>
                                         <%if Not(rsS.bof and rsS.eof) then
	rsS.move (Cint(NoncePage)-1)*50,1
	for i=1 to rsS.pagesize
%>
                                          <tr class="style10">
                                            <td height="22"><div align="center"><%=rsS("user_name")%></div></td>
                                            <td><div align="center"><%=rsS("user_number")%></div></td>
                                            <td><div align="center"><%=rsS("train_to")%></div></td>
                                            <td><div align="center"><%=rsS("user_yx")%></div></td>
											<td><div align="center"><%=rsS("user_pyxz")%></div></td>
                                            <td><div align="center"><%=rsS("train_info")%></div></td>
                                          </tr>
                                        
								 <%rsS.movenext
if rsS.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>没有找到任何记录!!!</font></marquee></td></tr>"
end if	
rsS.close
set rsS=nothing

								  %>
                                        </table></td>
                                      </tr>
                                      <tr>
									  <td>
									  <input type="hidden" name="page" value="<%=NoncePage%>">
                                <form name="form1" method="post" action="train_info.asp">
                                  <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                      <td valign="middle"><div align="right">
                                        <div align="right"> <span class="style2">
                                          <%
if NoncePage>1 then
	response.write "|<a href=train_info.asp?page=1>首 页</a>| |<a href=train_info.asp?page="&NoncePage-1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=train_info.asp?page="&NoncePage+1&">下一页</a>| |<a href=train_info.asp?page="&NumPage&">尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
  &nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录  &nbsp;&nbsp;转到
                                          <input name="page" type="text" class="style3" id="page" size="2">
                                          页</span>&nbsp;&nbsp; </div>
                                                  </div></td>
                                              </tr>
                                    </table>
                                          </form>
									  
									  
									  </td>
									  </tr>
                                      
                                     
                                    </table>
                                </div></td>
                              </tr>
                            </table>
</body>