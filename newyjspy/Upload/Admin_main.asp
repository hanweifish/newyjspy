<!--#include file="Config.asp"-->
<!--#include file="../admin/session.asp"-->
<%
Dim Msg
	If Request.QueryString("Action") = "Save" Then SaveData
	Sub SaveData()
	myConn.execute("update Config1 set OKAr='"&Request.Form("ftype")&"',OKsize="&Request.Form("fsize"))
	Msg = "�ɹ��޸����ļ�������Ϣ"
	End Sub

If msg <> "" Then
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_Main.asp'><link href='../style.css' rel='stylesheet' type='text/css'><div align='center'><br><br><br><br>"&Msg&"<br>��ҳ����3���ڷ���<BR>�����������û�з�Ӧ����<a href=Admin_Main.asp>����˴�����</a></div>")
Response.End()
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ϴ�����</title>
<link href="../style.css" rel="stylesheet" type="text/css">
</head>

<body>
<div align="center">
  <!--#include file="top1.asp" -->
  <table width="840" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="20" rowspan="2" background="../images/leftbk.jpg">&nbsp;</td>
      <td height="40" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;���ĵ�ǰλ�ã�&gt;&gt; <span class="style3"><a href="index.asp">�ϴ��ļ�</a></span> -- <a href="Admin_List.asp">�ļ�����</a> -- <span class="style3">ϵͳ����</span>&nbsp;&nbsp; <a href="../admin/admin_logout.asp">��<span class="style2">ע��</span>��</a></td>
      <td width="20" rowspan="2" background="../images/rightbk.jpg">&nbsp;</td>
    </tr>
    <tr>
      <td><div align="center">
          <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
            <form name="Edit" action="Admin_Main.asp?Action=Save" method="post">
              <tr>
                <td height="25"><%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from Config1"
	frst.open sql,myconn,1,1
	If not frst.Eof then
	%>
                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
                      <tr align="center" class="text">
                        <td height="10" colspan="2" bgcolor="#FFFFFF">&nbsp;</td>
                      </tr>
                      <tr align="center" class="text">
                        <td height="50" colspan="2" bgcolor="#FFFFFF">��̨����-ϵͳ���� [<a href="Admin_List.asp">��˽����ļ�����</a>]</td>
                      </tr>
                      <tr class="text">
                        <td width="200" height="50" bgcolor="#FFFFFF"><div align="right">�����ϴ����ļ����ͣ�</div></td>
                        <td bgcolor="#FFFFFF"><input name="ftype" type="text" class="style2" value="<%=rs(0)%>" size="50">
                        <br><br>����,�ָ���׺��,�м��������ϴ�asp/exe�ļ�</td>
                      </tr>
                      <tr class="text">
                        <td width="200" height="50" bgcolor="#FFFFFF"><div align="right">�����ϴ����ļ���С��</div></td>
                        <td bgcolor="#FFFFFF"><input name="fsize" type="text" class="style2" value="<%=rs(1)%>" size="15">
                        ��λ:Byte</td>
                      </tr>
                      <tr class="text">
                        <td height="1" colspan="2" align="right" bgcolor="#FFFFFF"><div align="center">
                          <input name="Submit" type="submit" class="style2" value="��  ��">
                        </div></td>
                      </tr>
                    </table>
                    <%
	  else
	  %>
                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
                      <tr class="text">
                        <td bgcolor="#FFFFFF">û�ж�Ӧ�����ݣ�</td>
                      </tr>
                    </table>
                    <%
	  end if
	  frst.close
	  set frst = nothing
	  myconn.close
	  set myconn = nothing
	  %>
                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
                      <tr class="text">
                        <td align="right" class="heading" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;</td>
                      </tr>
                  </table></td>
              </tr>
            </form>
          </table>
      </div></td>
    </tr>
  </table>
  <!--#include file="bottom1.asp" -->
</div>
</body>


</html>