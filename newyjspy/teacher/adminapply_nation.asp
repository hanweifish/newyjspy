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
sql="select * from apply_nation order by Pclb,user_tutor"
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
        <td height="24" width="40"><div align="center" class="style10">����</div></td>
        <td height="24"><div align="center" class="style10">ѧ��</div></td>
        <td height="24" align="center"><div align="center" class="style10">רҵ</div></td>
        <td height="24" width="40"><div align="center" class="style10">���</div></td>
        <td><div align="center" class="style10">��������</div></td>
        <td height="24" width="45"><div align="center" class="style10">��ʦ</div></td>
        <td><div align="center" class="style10">����ɼ�</div></td>
        <td><div align="center" class="style10">������ѧУרҵ</div></td>
        <td><div align="center" class="style10">�ɳ����</div></td>
        <td><div align="center" class="style10">�ƻ��ɳ�����</div></td>
        <td><div align="center" class="style10">��ϵ�绰</div></td>
        <td><div align="center" class="style10">E_mail</div></td>
        <td height="24" width="30"><div align="center" class="style10">ɾ��</div></td>
      </tr>
      <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*10,1
	for i=1 to rs.pagesize
%>
      <tr>
        <td height="24"><div align="center" class="style10"><a href="name_search.asp?user_number=<%=rs("user_number")%>" class="style3"><%=rs("user_name")%></a></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_number")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_major")%></div></td>
        <td height="24"><div align="center" class="style10"><%=rs("user_Pyxz")%></div></td>
        <td><div align="center" class="style10"><%=rs("user_Csrq")%></div></td>
        <td><div align="center" class="style10">&nbsp;<%=rs("user_Tutor")%></div></td>
        <td><div align="center" class="style10"><%=rs("sorts")%>,<%=rs("score")%></div></td>
        <td><div align="center" class="style10"><%=rs("school")%>,<%=rs("major")%></div></td>
        <td><div align="center" class="style10"><%=rs("Pclb")%></div></td>
        <td><div align="center" class="style10"><%=rs("applydate")%></div></td>
        <td><div align="center" class="style10"><%=rs("phone")%></div></td>
        <td><div align="center" class="style10"><%=rs("email")%></div></td>
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=adminapply_nationdel.asp?NoncePage=<%=NoncePage%>&ID=<%=rs("ID")%>><font color="#ff6633">ɾ��</font></a></div></td>
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
            <form action="adminapply_nation.asp" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=adminapply_nation.asp?page=1>�� ҳ</a>| |<a href=adminapply_nation.asp?page="&NoncePage-1&">��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=adminapply_nation.asp?page="&NoncePage+1&">��һҳ</a>| |<a href=adminapply_nation.asp?page="&NumPage&">β ҳ</a>|"
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
</div>
</body>
</html>
