<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("admin_account")="" or session("user_group")="subadmin" then
Response.write"�Բ�������û�е�½�����޴�Ȩ�ޣ�"
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
			alert("������ѧ��ѧ�ţ�");
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
<title>���ɳ����걨</title>
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
        <td height="24"><div align="center" class="style10">����</div></td>
        <td height="24"><div align="center" class="style10">ѧ��</div></td>
      
        <td><div align="center" class="style10">�ⶨ����ʱ��</div></td>
       
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
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>û���ҵ��κμ�¼!!!</font></marquee></td></tr>"
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
	response.write "|<a href=adminapplydate.asp?page=1>�� ҳ</a>| |<a href=adminapplydate.asp?page="&NoncePage-1&">��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=adminapplydate.asp?page="&NoncePage+1&">��һҳ</a>| |<a href=adminapplydate.asp?page="&NumPage&">β ҳ</a>|"
else
	response.write "|��һҳ| |β ҳ|"
end if
%>
                        &nbsp;ҳ�Σ�<font color="#0033CC" class="style3"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> ��<font color="#0033CC" class="style3"><%=NumRecord%></font>����¼  &nbsp;&nbsp;ת��
                        <input name="page" type="text" class="style3" id="page" size="2" />
                        ҳ</span>&nbsp;&nbsp; </div>
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

    <p>�����������ѧ����email:<br>
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
