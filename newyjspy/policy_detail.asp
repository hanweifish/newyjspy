<!--#include file="conn.asp"-->
<%
dim policy_ID
policy_ID=trim(request("policy_ID"))
set rsn=server.createobject("adodb.recordset")
sql="select * from policy where policy_ID="&policy_ID
rsn.open sql,conn,1,1
%>

<%
function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
    HTMLEncode = fString
end if
end function
%>

<html>
<head>
<script language="javascript">
<!--

window.status="��ӭ�����Ͼ���ѧ����ϵ�о���������Ϣϵͳ��"
//-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�о�����Ϣ����ϵͳ</title>
<link href="style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style9 {font-size: 12px}
-->
</style>
<style type="text/css">
<!--
.style15 {font-size: 14px}
-->
</style>
</head>
<body>
<div align="center">
  <table width="603"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="54" background="user/userimages/titlebk1.gif"><div align="center"><img src="user/userimages/policy.gif"></div></td>
                    </tr>
                    <tr>
                      <td height="53" background="user/userimages/titlebk2.gif">&nbsp;</td>
                    </tr>
                    <tr>
                      <td background="user/userimages/titlebk.gif"><div align="center">
                        <table width="80%"  border="0" cellpadding="0" cellspacing="0" class="thin">
                          <tr>
                            <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td height="25" ><div align="center" class="style3"><%=rsn("policy_title")%></div></td>
                                </tr>
                                <tr>
                                  <td><div align="center">
                                    <table width="80%"  border="0" cellspacing="0" cellpadding="0">
                                        <tr>
                                          <td width="39%" height="20" class="style9">�����ˣ�<%=rsn("policy_author")%></td>
                                          <td width="17%" class="style9"><div align="right"></div></td>
                                          <td width="44%" class="style9"><div align="right">����ʱ�䣺<%=rsn("policy_time")%>&nbsp;&nbsp;</div></td>
                                        </tr>
                                        <tr>
                                          <td height="20" colspan="3"><div align="left"></div></td>
                                        </tr>
                                        <tr>
                                          <td colspan="3"><br>
                                              <span class="style2"><%=HTMLEncode(rsn("policy_content"))%></span><br>
&nbsp;</td>
                                        </tr>
                                        <tr class="style2">
                                          <td height="25" colspan="3">��ע��Ϣ��</td>
                                      </tr>
                                        <tr class="style2">
                                          <td colspan="3"><%=HTMLEncode(rsn("policy_info"))%></td>
                                      </tr>
                                                                    </table>
                                  </div></td>
                                </tr>
                            </table></td>
                          </tr>
                        </table>
                        </div></td>
                    </tr>
                    <tr>
                      <td height="34" background="user/userimages/titlebk3.gif">&nbsp;</td>
                    </tr>
                    </table>
</div>
</body>
</html>
