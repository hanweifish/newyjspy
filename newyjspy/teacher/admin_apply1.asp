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
sorts=request("sorts")

%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�������鿴</title>
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
        <td><div align="center" class="style10">����</div></td>
        <td><div align="center" class="style10">ѧ��</div></td>
        <td><div align="center" class="style10">רҵ</div></td>
        <td><div align="center" class="style10">��������</div></td>
        <td><div align="center" class="style10">�Ա�</div></td>
        <td><div align="center" class="style10">Ժϵ</div></td>
        
        <td><div align="center" class="style10">��������ʱ��ҵ</div></td>
        <td><div align="center" class="style10">����ʱ��</div></td>
        <td><div align="center" class="style10">��ϵ�绰</div></td>
   
        <td><div align="center" class="style10">ɾ��</div></td>
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
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">ɾ��</font></a></div></td>
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
            <form action="admin_apply1.asp?sorts=yqby" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=yqby>�� ҳ</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=yqby>��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=yqby>��һҳ</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=yqby>β ҳ</a>|"
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
        <td><div align="center" class="style10">����</div></td>
        <td><div align="center" class="style10">ѧ��</div></td>
        <td><div align="center" class="style10">רҵ</div></td>
        <td><div align="center" class="style10">��������</div></td>
        <td><div align="center" class="style10">�Ա�</div></td>
        <td><div align="center" class="style10">Ժϵ</div></td>
        
        <td><div align="center" class="style10">��ǰ����ʱ��ҵ</div></td>
        <td><div align="center" class="style10">����ʱ��</div></td>
        <td><div align="center" class="style10">��ϵ�绰</div></td>
   
        <td><div align="center" class="style10">ɾ��</div></td>
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
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">ɾ��</font></a></div></td>
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
            <form action="admin_apply1.asp?sorts=tqby" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=tqby>�� ҳ</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=tqby>��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=tqby>��һҳ</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=tqby>β ҳ</a>|"
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
        <td><div align="center" class="style10">����</div></td>
        <td><div align="center" class="style10">ѧ��</div></td>
        <td><div align="center" class="style10">רҵ</div></td>
        <td><div align="center" class="style10">��������</div></td>
        <td><div align="center" class="style10">�Ա�</div></td>
        <td><div align="center" class="style10">Ժϵ</div></td>
        
        <td><div align="center" class="style10">��ѧԭ��</div></td>
        <td><div align="center" class="style10">����ʱ��</div></td>
        <td><div align="center" class="style10">��ϵ�绰</div></td>
   
        <td><div align="center" class="style10">ɾ��</div></td>
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
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">ɾ��</font></a></div></td>
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
            <form action="admin_apply1.asp?sorts=tx" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=tx>�� ҳ</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=tx>��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=tx>��һҳ</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=tx>β ҳ</a>|"
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
        <td><div align="center" class="style10">����</div></td>
        <td><div align="center" class="style10">ѧ��</div></td>
        <td><div align="center" class="style10">רҵ</div></td>
        <td><div align="center" class="style10">��������</div></td>
        <td><div align="center" class="style10">�Ա�</div></td>
        <td><div align="center" class="style10">Ժϵ</div></td>
        
        <td><div align="center" class="style10">��ѧԭ��</div></td>
        <td><div align="center" class="style10">����ʱ��</div></td>
        <td><div align="center" class="style10">��ϵ�绰</div></td>
   
        <td><div align="center" class="style10">ɾ��</div></td>
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
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">ɾ��</font></a></div></td>
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
            <form action="admin_apply1.asp?sorts=xx" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=xx>�� ҳ</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=xx>��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=xx>��һҳ</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=xx>β ҳ</a>|"
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
        <td><div align="center" class="style10">����</div></td>
        <td><div align="center" class="style10">ѧ��</div></td>
        <td><div align="center" class="style10">רҵ</div></td>
        <td><div align="center" class="style10">��������</div></td>
        <td><div align="center" class="style10">�Ա�</div></td>
        <td><div align="center" class="style10">Ժϵ</div></td>
        
        <td><div align="center" class="style10">��ѧ�ڼ�</div></td>
        <td class="style10"><div align="center">��ѧ����Ƿ�ͨ��</div></td>
        <td><div align="center" class="style10">����ʱ��</div></td>
        <td><div align="center" class="style10">��ϵ�绰</div></td>
   
        <td><div align="center" class="style10">ɾ��</div></td>
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
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">ɾ��</font></a></div></td>
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
        <td height="40" colspan="15"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=fx" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=fx>�� ҳ</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=fx>��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=fx>��һҳ</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=fx>β ҳ</a>|"
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
        <td><div align="center" class="style10">����</div></td>
        <td><div align="center" class="style10">ѧ��</div></td>
        <td><div align="center" class="style10">רҵ</div></td>
        <td><div align="center" class="style10">��������</div></td>
        <td><div align="center" class="style10">�Ա�</div></td>
        <td><div align="center" class="style10">Ժϵ</div></td>
        
        <td><div align="center" class="style10">��תרҵ</div></td>
        <td class="style10"><div align="center">��ʦ����</div></td>
        <td class="style10"><div align="center">��תԺϵ</div></td>
        <td class="style10"><div align="center">תרҵԭ��</div></td>		
        <td><div align="center" class="style10">����ʱ��</div></td>
        <td><div align="center" class="style10">��ϵ�绰</div></td>
   
        <td><div align="center" class="style10">ɾ��</div></td>
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
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">ɾ��</font></a></div></td>
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
        <td height="40" colspan="16"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=zzy" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=zzy>�� ҳ</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=zzy>��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=zzy>��һҳ</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=zzy>β ҳ</a>|"
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
        <td><div align="center" class="style10">����</div></td>
        <td><div align="center" class="style10">ѧ��</div></td>
        <td><div align="center" class="style10">רҵ</div></td>
        <td><div align="center" class="style10">��������</div></td>
        <td><div align="center" class="style10">�Ա�</div></td>
        <td><div align="center" class="style10">Ժϵ</div></td>
        
        <td><div align="center" class="style10">��ת��ʦ����</div></td>
        <td><div align="center" class="style10">ԭ��ʦ����</div></td>		
        <td class="style10"><div align="center">ת��ʦԭ��</div></td>
        
        <td><div align="center" class="style10">����ʱ��</div></td>
        <td><div align="center" class="style10">��ϵ�绰</div></td>
   
        <td><div align="center" class="style10">ɾ��</div></td>
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
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">ɾ��</font></a></div></td>
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
        <td height="40" colspan="16"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=zds" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=zds>�� ҳ</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=zds>��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=zds>��һҳ</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=zds>β ҳ</a>|"
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
        <td><div align="center" class="style10">����</div></td>
        <td><div align="center" class="style10">ѧ��</div></td>
        <td><div align="center" class="style10">רҵ</div></td>
        <td><div align="center" class="style10">��������</div></td>
        <td><div align="center" class="style10">�Ա�</div></td>
        <td><div align="center" class="style10">Ժϵ</div></td>
        
        <td><div align="center" class="style10">ȡ��ѧ��ԭ��</div></td>
        
        <td><div align="center" class="style10">����ʱ��</div></td>
        <td><div align="center" class="style10">��ϵ�绰</div></td>
   
        <td><div align="center" class="style10">ɾ��</div></td>
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
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">ɾ��</font></a></div></td>
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
        <td height="40" colspan="16"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=qxxj" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=qxxj>�� ҳ</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=qxxj>��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=qxxj>��һҳ</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=qxxj>β ҳ</a>|"
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
        <td><div align="center" class="style10">����</div></td>
        <td><div align="center" class="style10">ѧ��</div></td>
        <td><div align="center" class="style10">רҵ</div></td>
        <td><div align="center" class="style10">��������</div></td>
        <td><div align="center" class="style10">�Ա�</div></td>
        <td><div align="center" class="style10">Ժϵ</div></td>
        
        <td><div align="center" class="style10">˶ʿѧ��</div></td>
        <td><div align="center" class="style10">˶ʿרҵ</div></td>
        <td class="style10"><div align="center">��ʦ����</div></td>
        <td class="style10"><div align="center">��ֹԭ��</div></td>
        <td><div align="center" class="style10">����ʱ��</div></td>
        <td><div align="center" class="style10">��ϵ�绰</div></td>
   
        <td><div align="center" class="style10">ɾ��</div></td>
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
   
        
        <td height="24"><div align="center" class="style3 STYLE11"><a href=admin_applydel.asp><font color="#ff6633">ɾ��</font></a></div></td>
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
        <td height="40" colspan="16"><div align="center">
            <input type="hidden" name="page" value="<%=NoncePage%>" />
            <form action="admin_apply1.asp?sorts=zzsbld" method="post" name="form1" id="form1">
              <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle"><div align="right">
                      <div align="right"> <span class="style10">
                        <%
if NoncePage>1 then
	response.write "|<a href=admin_apply1.asp?page=1&sorts=zzsbld>�� ҳ</a>| |<a href=admin_apply1.asp?page="&NoncePage-1&"&sorts=zzsbld>��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=admin_apply1.asp?page="&NoncePage+1&"&sorts=zzsbld>��һҳ</a>| |<a href=admin_apply1.asp?page="&NumPage&"&sorts=zzsbld>β ҳ</a>|"
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
end if
%>	



</div>
</body>
</html>
