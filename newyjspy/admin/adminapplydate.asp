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
set rs=server.createobject("adodb.recordset")
sql="select user_info.user_name,user_info.user_number,apply.apartdate,apply.apply_id from apply inner join user_info on user_info.user_number=apply.user_number order by apply.apply_id "
rs.open sql,conn,1,1
%>

<script language="javascript">
	function checkform(form3)
	{
		if (document.form3.studentno.value=="")
		{
			alert("请输入学生学号！");
		}
		else
		{
			form3.submit();
		}
		return false;
	}
</script>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>公派出国申报</title>
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
    <%       if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=20
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
      <tr>
        <td height="24"><div align="center" class="style10">姓名</div></td>
        <td height="24"><div align="center" class="style10">学号</div></td>
      
        <td><div align="center" class="style10">拟定出发时间</div></td>
       
      </tr>
      <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*20,1
	for i=1 to rs.pagesize
%>
      <tr>
        <td height="24"><div align="center" class="style10"><a href="number_search.asp?user_number=<%=rs("user_number")%>" class="style3"><%=rs("user_name")%></a></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_number")%></div></td>

        <td><div align="center" class="style10">&nbsp;<%=rs("apartdate")%></div></td>
        
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
            <form action="adminapplydate.asp" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=adminapplydate.asp?page=1>首 页</a>| |<a href=adminapplydate.asp?page="&NoncePage-1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=adminapplydate.asp?page="&NoncePage+1&">下一页</a>| |<a href=adminapplydate.asp?page="&NumPage&">尾 页</a>|"
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
set rs1=server.createobject("adodb.recordset")
sql="select apply_email from apply order by apply_id "
rs1.open sql,conn,1,1
%>

    <p>输出所有申请学生的email:<br>
<% 
for i=1 to rs1.recordcount
%>	
 <%=rs1("apply_email")%> ;   
	  <%rs1.movenext
if rs1.eof then exit for
next

rs1.close
set rs1=nothing
%>
	</p>
</div>
</body>
</html>
