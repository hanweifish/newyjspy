<!--#include file ="conn.asp"-->
<!--#include file="../admin/session.asp"-->

<%
	dim forum_ID,page
	page=trim(request("page"))
	forum_ID=trim(request("forum_ID"))
%>

<%
	set rs = Server.createobject("adodb.recordset")
	sql = "select * from forum where forum_ID="&forum_ID
	rs.open sql,conn,3,3
	set rs1 = Server.createobject("adodb.recordset")
	sql = "select * from reforum where forum_ID="&forum_ID
	rs1.open sql,conn,3,3
%>

<%
	if not (rs1.bof and rs1.eof) then
	for i=1 to rs1.recordcount
	rs1.delete
	rs1.movenext
	if rs1.eof and rs1.bof then exit for
	next
	end if
	rs1.close
	set rs1=nothing
	rs.delete
	rs.close
	set rs=nothing
%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="refresh" content="3;URL=admin_index.asp?page=<%=page%>">
<title>��������</title>
<link href="../style.css" rel="stylesheet" type="text/css">
</head>

<body >
<div align="center">
<!--#include file = "top2.asp"-->
<table width="800"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="center">
      <p>&nbsp;</p>
      <p><span>ɾ�����Գɹ���<br>
          ��ҳ����3��󷵻�<br>
      ������������û�з�Ӧ����<a href=admin_index.asp?page=<%=page%>><b>����˴�����</b></a></span>
&nbsp;</p>
    </div></td>
  </tr>
</table>
</div>
</body>
</html>
