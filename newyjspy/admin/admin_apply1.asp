<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("admin_account")="" or session("user_group")="subadmin" then
Response.write"对不起，您还没有登陆或者无此权限！"
Response.end
end if
%>

<%
dim admin_account
admin_account=session("admin_account")
sorts=request("sorts")

%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>申请结果查看</title>
<style type="text/css">
<!--
.style10 {font-size: 12px;
	color: #004080;
}
.STYLE11 {font-size: 10pt}
-->
</style>
</head>

<body>
<div align="center" class="style10"><br />
   
<%
if sorts="yqby" then
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_yqby='y'"
rs.open sql,conn,1,1
%>          

<%       if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=30
			NumPage=rs.Pagecount
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
<table width="100%" height="145"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
      <tr height="30">
        <td><div align="center" class="style10">姓名</div></td>
        <td><div align="center" class="style10">学号</div></td>
        <td><div align="center" class="style10">专业</div></td>
        <td><div align="center" class="style10">培养性质</div></td>
        <td><div align="center" class="style10">性别</div></td>
        <td><div align="center" class="style10">院系</div></td>
        
        <td><div align="center" class="style10">延期至何时毕业</div></td>
        <td><div align="center" class="style10">申请时间</div></td>
        <td><div align="center" class="style10">联系电话</div></td>
   
        <td><div align="center" class="style10">删除</div></td>
      </tr>
      <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*30,1
	for i=1 to rs.pagesize
%>
      <tr>
        <td height="24"><div align="center" class="style10"><a href="name_search.asp?user_number=<%=rs("user_number")%>" class="style3"><%=rs("user_name")%></a></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_number")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_major")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_Pyxz")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sex")%></div></td>
        <td><div align="center" class="style10">&nbsp;<%=rs("user_yx")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_yqby1")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_yqbysq")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sqphone")%></div></td>
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">删除</font></a></div></td>
      </tr>
      <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>没有找到任何记录!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
      <tr>
        <td height="40" colspan="14"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=yqby" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=yqby>首 页</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=yqby>上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=yqby>下一页</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=yqby>尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
                        &nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录  &nbsp;&nbsp;转到
                        <input name="page" type="text" class="style3" id="page" size="2" />
                        页</span>&nbsp;&nbsp; </div>
                  </div></td>
                </tr>
              </table>
            </form>
        </div></td>
      </tr>
  </table>
<%
end if
%>	

<%
if sorts="tqby" then
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_tqby='y'"
rs.open sql,conn,1,1
%>          

<%       if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=30
			NumPage=rs.Pagecount
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
<table width="100%" height="145"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
      <tr height="30">
        <td><div align="center" class="style10">姓名</div></td>
        <td><div align="center" class="style10">学号</div></td>
        <td><div align="center" class="style10">专业</div></td>
        <td><div align="center" class="style10">培养性质</div></td>
        <td><div align="center" class="style10">性别</div></td>
        <td><div align="center" class="style10">院系</div></td>
        
        <td><div align="center" class="style10">提前至何时毕业</div></td>
        <td><div align="center" class="style10">申请时间</div></td>
        <td><div align="center" class="style10">联系电话</div></td>
   
        <td><div align="center" class="style10">删除</div></td>
      </tr>
      <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*30,1
	for i=1 to rs.pagesize
%>
      <tr>
        <td height="24"><div align="center" class="style10"><a href="name_search.asp?user_number=<%=rs("user_number")%>" class="style3"><%=rs("user_name")%></a></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_number")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_major")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_Pyxz")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sex")%></div></td>
        <td><div align="center" class="style10">&nbsp;<%=rs("user_yx")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_tqby1")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_tqbysq")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sqphone")%></div></td>
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">删除</font></a></div></td>
      </tr>
      <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>没有找到任何记录!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
      <tr>
        <td height="40" colspan="14"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=tqby" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=tqby>首 页</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=tqby>上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=tqby>下一页</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=tqby>尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
                        &nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录  &nbsp;&nbsp;转到
                        <input name="page" type="text" class="style3" id="page" size="2" />
                        页</span>&nbsp;&nbsp; </div>
                  </div></td>
                </tr>
              </table>
            </form>
        </div></td>
      </tr>
  </table>
<%
end if
%>	

<%
if sorts="tx" then
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_tx='y'"
rs.open sql,conn,1,1
%>          

<%       if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=30
			NumPage=rs.Pagecount
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
<table width="100%" height="145"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
      <tr height="30">
        <td><div align="center" class="style10">姓名</div></td>
        <td><div align="center" class="style10">学号</div></td>
        <td><div align="center" class="style10">专业</div></td>
        <td><div align="center" class="style10">培养性质</div></td>
        <td><div align="center" class="style10">性别</div></td>
        <td><div align="center" class="style10">院系</div></td>
        
        <td><div align="center" class="style10">退学原因</div></td>
        <td><div align="center" class="style10">申请时间</div></td>
        <td><div align="center" class="style10">联系电话</div></td>
   
        <td><div align="center" class="style10">删除</div></td>
      </tr>
      <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*30,1
	for i=1 to rs.pagesize
%>
      <tr>
        <td height="24"><div align="center" class="style10"><a href="name_search.asp?user_number=<%=rs("user_number")%>" class="style3"><%=rs("user_name")%></a></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_number")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_major")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_Pyxz")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sex")%></div></td>
        <td><div align="center" class="style10">&nbsp;<%=rs("user_yx")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_tx1")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_txsq")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sqphone")%></div></td>
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">删除</font></a></div></td>
      </tr>
      <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>没有找到任何记录!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
      <tr>
        <td height="40" colspan="14"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=tx" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=tx>首 页</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=tx>上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=tx>下一页</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=tx>尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
                        &nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录  &nbsp;&nbsp;转到
                        <input name="page" type="text" class="style3" id="page" size="2" />
                        页</span>&nbsp;&nbsp; </div>
                  </div></td>
                </tr>
              </table>
            </form>
        </div></td>
      </tr>
  </table>
<%
end if
%>	

<%
if sorts="xx" then
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_xx='y'"
rs.open sql,conn,1,1
%>          

<%       if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=30
			NumPage=rs.Pagecount
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
<table width="100%" height="145"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
      <tr height="30">
        <td><div align="center" class="style10">姓名</div></td>
        <td><div align="center" class="style10">学号</div></td>
        <td><div align="center" class="style10">专业</div></td>
        <td><div align="center" class="style10">培养性质</div></td>
        <td><div align="center" class="style10">性别</div></td>
        <td><div align="center" class="style10">院系</div></td>
        
        <td><div align="center" class="style10">休学原因</div></td>
        <td><div align="center" class="style10">申请时间</div></td>
        <td><div align="center" class="style10">联系电话</div></td>
   
        <td><div align="center" class="style10">删除</div></td>
      </tr>
      <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*30,1
	for i=1 to rs.pagesize
%>
      <tr>
        <td height="24"><div align="center" class="style10"><a href="name_search.asp?user_number=<%=rs("user_number")%>" class="style3"><%=rs("user_name")%></a></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_number")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_major")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_Pyxz")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sex")%></div></td>
        <td><div align="center" class="style10">&nbsp;<%=rs("user_yx")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_xx1")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_xxsq")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sqphone")%></div></td>
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">删除</font></a></div></td>
      </tr>
      <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>没有找到任何记录!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
      <tr>
        <td height="40" colspan="14"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=xx" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=xx>首 页</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=xx>上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=xx>下一页</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=xx>尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
                        &nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录  &nbsp;&nbsp;转到
                        <input name="page" type="text" class="style3" id="page" size="2" />
                        页</span>&nbsp;&nbsp; </div>
                  </div></td>
                </tr>
              </table>
            </form>
        </div></td>
      </tr>
  </table>
<%
end if
%>	

<%
if sorts="fx" then
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_fx='y'"
rs.open sql,conn,1,1
%>          

<%       if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=30
			NumPage=rs.Pagecount
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
<table width="100%" height="145"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
      <tr height="30">
        <td><div align="center" class="style10">姓名</div></td>
        <td><div align="center" class="style10">学号</div></td>
        <td><div align="center" class="style10">专业</div></td>
        <td><div align="center" class="style10">培养性质</div></td>
        <td><div align="center" class="style10">性别</div></td>
        <td><div align="center" class="style10">院系</div></td>
        
        <td><div align="center" class="style10">休学期间</div></td>
        <td class="style10"><div align="center">复学审查是否通过</div></td>
        <td><div align="center" class="style10">申请时间</div></td>
        <td><div align="center" class="style10">联系电话</div></td>
   
        <td><div align="center" class="style10">删除</div></td>
      </tr>
      <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*30,1
	for i=1 to rs.pagesize
%>
      <tr>
        <td height="24"><div align="center" class="style10"><a href="name_search.asp?user_number=<%=rs("user_number")%>" class="style3"><%=rs("user_name")%></a></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_number")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_major")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_Pyxz")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sex")%></div></td>
        <td><div align="center" class="style10">&nbsp;<%=rs("user_yx")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_fx1")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_fx2")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_fxsq")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sqphone")%></div></td>
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">删除</font></a></div></td>
      </tr>
      <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>没有找到任何记录!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
      <tr>
        <td height="40" colspan="15"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=fx" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=fx>首 页</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=fx>上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=fx>下一页</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=fx>尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
                        &nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录  &nbsp;&nbsp;转到
                        <input name="page" type="text" class="style3" id="page" size="2" />
                        页</span>&nbsp;&nbsp; </div>
                  </div></td>
                </tr>
              </table>
            </form>
        </div></td>
      </tr>
    </table>
<%
end if
%>	

<%
if sorts="zzy" then
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_zzy='y'"
rs.open sql,conn,1,1
%>          

<%       if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=30
			NumPage=rs.Pagecount
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
<table width="100%" height="145"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
      <tr height="30">
        <td><div align="center" class="style10">姓名</div></td>
        <td><div align="center" class="style10">学号</div></td>
        <td><div align="center" class="style10">专业</div></td>
        <td><div align="center" class="style10">培养性质</div></td>
        <td><div align="center" class="style10">性别</div></td>
        <td><div align="center" class="style10">院系</div></td>
        
        <td><div align="center" class="style10">拟转专业</div></td>
        <td class="style10"><div align="center">导师姓名</div></td>
        <td class="style10"><div align="center">拟转院系</div></td>
        <td class="style10"><div align="center">转专业原因</div></td>		
        <td><div align="center" class="style10">申请时间</div></td>
        <td><div align="center" class="style10">联系电话</div></td>
   
        <td><div align="center" class="style10">删除</div></td>
      </tr>
      <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*30,1
	for i=1 to rs.pagesize
%>
      <tr>
        <td height="24"><div align="center" class="style10"><a href="name_search.asp?user_number=<%=rs("user_number")%>" class="style3"><%=rs("user_name")%></a></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_number")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_major")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_Pyxz")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sex")%></div></td>
        <td><div align="center" class="style10">&nbsp;<%=rs("user_yx")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_zzy1")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_zzy3")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_zzy2")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_zzy4")%></div></td>		
        <td><div align="center" class="style10"><%=rs("user_zzysq")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sqphone")%></div></td>
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">删除</font></a></div></td>
      </tr>
      <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>没有找到任何记录!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
      <tr>
        <td height="40" colspan="16"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=zzy" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=zzy>首 页</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=zzy>上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=zzy>下一页</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=zzy>尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
                        &nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录  &nbsp;&nbsp;转到
                        <input name="page" type="text" class="style3" id="page" size="2" />
                        页</span>&nbsp;&nbsp; </div>
                  </div></td>
                </tr>
              </table>
            </form>
        </div></td>
      </tr>
    </table>
<%
end if
%>	

<%
if sorts="zds" then
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_zds='y'"
rs.open sql,conn,1,1
%>          

<%       if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=30
			NumPage=rs.Pagecount
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
<table width="100%" height="145"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
      <tr height="30">
        <td><div align="center" class="style10">姓名</div></td>
        <td><div align="center" class="style10">学号</div></td>
        <td><div align="center" class="style10">专业</div></td>
        <td><div align="center" class="style10">培养性质</div></td>
        <td><div align="center" class="style10">性别</div></td>
        <td><div align="center" class="style10">院系</div></td>
        
        <td><div align="center" class="style10">拟转导师姓名</div></td>
        <td><div align="center" class="style10">原导师姓名</div></td>		
        <td class="style10"><div align="center">转导师原因</div></td>
        
        <td><div align="center" class="style10">申请时间</div></td>
        <td><div align="center" class="style10">联系电话</div></td>
   
        <td><div align="center" class="style10">删除</div></td>
      </tr>
      <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*30,1
	for i=1 to rs.pagesize
%>
      <tr>
        <td height="24"><div align="center" class="style10"><a href="name_search.asp?user_number=<%=rs("user_number")%>" class="style3"><%=rs("user_name")%></a></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_number")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_major")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_Pyxz")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sex")%></div></td>
        <td><div align="center" class="style10">&nbsp;<%=rs("user_yx")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_zds1")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_zds3")%></div></td>		
        <td><div align="center" class="style10"><%=rs("user_zds2")%></div></td>

        <td><div align="center" class="style10"><%=rs("user_zdssq")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sqphone")%></div></td>
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">删除</font></a></div></td>
      </tr>
      <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>没有找到任何记录!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
      <tr>
        <td height="40" colspan="16"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=zds" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=zds>首 页</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=zds>上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=zds>下一页</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=zds>尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
                        &nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录  &nbsp;&nbsp;转到
                        <input name="page" type="text" class="style3" id="page" size="2" />
                        页</span>&nbsp;&nbsp; </div>
                  </div></td>
                </tr>
              </table>
            </form>
        </div></td>
      </tr>
    </table>
<%
end if
%>	

<%
if sorts="qxxj" then
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_qxxj='y'"
rs.open sql,conn,1,1
%>          

<%       if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=30
			NumPage=rs.Pagecount
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
<table width="100%" height="145"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
      <tr height="30">
        <td><div align="center" class="style10">姓名</div></td>
        <td><div align="center" class="style10">学号</div></td>
        <td><div align="center" class="style10">专业</div></td>
        <td><div align="center" class="style10">培养性质</div></td>
        <td><div align="center" class="style10">性别</div></td>
        <td><div align="center" class="style10">院系</div></td>
        
        <td><div align="center" class="style10">取消学籍原因</div></td>
        
        <td><div align="center" class="style10">申请时间</div></td>
        <td><div align="center" class="style10">联系电话</div></td>
   
        <td><div align="center" class="style10">删除</div></td>
      </tr>
      <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*30,1
	for i=1 to rs.pagesize
%>
      <tr>
        <td height="24"><div align="center" class="style10"><a href="name_search.asp?user_number=<%=rs("user_number")%>" class="style3"><%=rs("user_name")%></a></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_number")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_major")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_Pyxz")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sex")%></div></td>
        <td><div align="center" class="style10">&nbsp;<%=rs("user_yx")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_qxxj1")%></div></td>


        <td><div align="center" class="style10"><%=rs("user_qxxjsq")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sqphone")%></div></td>
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">删除</font></a></div></td>
      </tr>
      <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>没有找到任何记录!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
      <tr>
        <td height="40" colspan="16"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=qxxj" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=qxxj>首 页</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=qxxj>上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=qxxj>下一页</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=qxxj>尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
                        &nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录  &nbsp;&nbsp;转到
                        <input name="page" type="text" class="style3" id="page" size="2" />
                        页</span>&nbsp;&nbsp; </div>
                  </div></td>
                </tr>
              </table>
            </form>
        </div></td>
      </tr>
    </table>
<%
end if
%>	

<%
if sorts="zzsbld" then
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_zzsbld='y'"
rs.open sql,conn,1,1
%>          

<%       if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=30
			NumPage=rs.Pagecount
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
<table width="100%" height="145"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
      <tr height="30">
        <td><div align="center" class="style10">姓名</div></td>
        <td><div align="center" class="style10">学号</div></td>
        <td><div align="center" class="style10">专业</div></td>
        <td><div align="center" class="style10">培养性质</div></td>
        <td><div align="center" class="style10">性别</div></td>
        <td><div align="center" class="style10">院系</div></td>
        
        <td><div align="center" class="style10">硕士学号</div></td>
        <td><div align="center" class="style10">硕士专业</div></td>
        <td class="style10"><div align="center">导师姓名</div></td>
        <td class="style10"><div align="center">终止原因</div></td>
        <td><div align="center" class="style10">申请时间</div></td>
        <td><div align="center" class="style10">联系电话</div></td>
   
        <td><div align="center" class="style10">删除</div></td>
      </tr>
      <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*30,1
	for i=1 to rs.pagesize
%>
      <tr>
        <td height="24"><div align="center" class="style10"><a href="name_search.asp?user_number=<%=rs("user_number")%>" class="style3"><%=rs("user_name")%></a></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_number")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_major")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_Pyxz")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sex")%></div></td>
        <td><div align="center" class="style10">&nbsp;<%=rs("user_yx")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_zzsbld1")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_zzsbld2")%></div></td>		
        <td><div align="center" class="style10"><%=rs("user_zzsbld3")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_zzsbld4")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_zzsbldsq")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_sqphone")%></div></td>
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">删除</font></a></div></td>
      </tr>
      <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>没有找到任何记录!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
      <tr>
        <td height="40" colspan="16"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=zzsbld" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=zzsbld>首 页</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=zzsbld>上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=zzsbld>下一页</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=zzsbld>尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
                        &nbsp;页次：<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> 共<font color="#0033CC" class="style3"><%=NumRecord%></font>条记录  &nbsp;&nbsp;转到
                        <input name="page" type="text" class="style3" id="page" size="2" />
                        页</span>&nbsp;&nbsp; </div>
                  </div></td>
                </tr>
              </table>
            </form>
        </div></td>
      </tr>
    </table>
<%
end if
%>	



</div>
</body>
</html>
