<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("user_account")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
dim user_account,user_number
user_account=session("user_account")
user_number=session("user_number")
set rs=server.createobject("adodb.recordset")
sql="select * from user_info where user_account='"&user_account&"'"
rs.open sql,conn,1,1
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

<script language="javascript">
	function checkuser()
	{
		var msg = "";
		if(document.form.user_account.value.length < 6)
			{
				msg = msg + "     �������û���!\n\n";
			}
		if(document.form.user_pwd.value.length == "")
			{
				msg = msg + "   ������6-12λ����!\n\n";
			}
		if(document.form.user_pwd.value != ""&&(document.form.user_pwd.value.length < 6 || document.form.user_pwd.value.length > 12))
			{
				msg = msg + "  ���볤��Ҫ����6С��12!\n\n";
			}
		if(document.form.user_pwd1.value == "")
			{
				msg = msg + "    ������ȷ������!\n\n";
			}
		if(document.form.user_pwd1.value != ""&&(document.form.user_pwd1.value.length < 6 || document.form.user_pwd1.value.length > 16))
			{
				msg = msg + "ȷ�����볤��Ҫ����6С��12!\n\n";
			}
		if(document.form.user_pwd.value != document.form.user_pwd1.value)
			{
				msg = msg + " ������������벻ƥ��!\n\n";
			}
		if(document.form.user_roomphone.value.length < 8||document.form.user_roomphone.value == "�����")
			{
				msg = msg + "    ��������ϵ��ʽ!\n\n";
			}
		if(!/^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/.test(document.form.user_mail.value)||document.form.user_mail.value =="�����")
			{
				msg = msg + "    ����ȷ��������!\n\n";
			}
		if(document.form.user_address.value == ""||document.form.user_address.value == "�����")
			{
				msg = msg + "      �������ַ!\n\n";
			}
		if(document.form.user_code.value.length != 6||document.form.user_code.value == "�����")
			{
				msg = msg + "      �������ʱ�!\n\n";
			}
		if(msg !="") 
			{
				alert(msg);
				return false;
			}
		document.form.submit();
	}
</script>

<html>
<head>
<script language="javascript">
<!--

window.status="��ӭ�����Ͼ���ѧ����ϵ�о���������Ϣϵͳ��"
//-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�о�����Ϣ����ϵͳ</title>
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
.style13 {color: #FF0000}
-->
</style>
<!--#include file="top1.asp"-->
<table width="800"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="187" height="12"></td>
    <td></td>
  </tr>
  <tr>
    <td rowspan="3" valign="top"><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="47" background="../indeximages/stulogin.gif">&nbsp;</td>
      </tr>
      <tr>
        <td><div align="center">
          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="100"><div align="center">
                <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" background="../indeximages/loginbk.gif">
                  <tr>
                    <td width="20%"><div align="center">                      </div>
                         <div align="center"></div></td>
                    <td width="60%" class="style3"><%=rs("user_account")%>��</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td ><div align="center">                      </div>
                         <div align="center"></div></td>
                    <td height="30" class="style2">���Ѿ�<span class="style2">��¼�ɹ�</span>,��</td>
                         <td>&nbsp;</td>
                   </tr>
                  <tr>
                    <td><div align="center"></div></td>
                    <td height="30" class="style2">ѡ������Ҫ�ķ���!</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td height="10" colspan="2"><div align="center"></div></td>
                    <td width="20%">&nbsp;</td>
                  </tr>
                </table>
                </div></td>
            </tr>
            <tr>
              <td height="6" background="../indeximages/loginbk.gif"><div align="center"><img src="../indeximages/loginbar.gif" width="129" height="2"></div></td>
            </tr>
            <tr>
              <td height="60" valign="center" background="../indeximages/loginbk.gif"><div align="center"><a href="user_logout.asp"><img src="../includeimages/logout.gif" width="60" height="24" border="0"></a></div></td>
            </tr>
          </table>
        </div></td>
      </tr>
      <tr>
        <td height="77" background="../indeximages/links.gif">&nbsp;</td>
      </tr>
      <tr>
        <td background="../indeximages/loginbk.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center">
                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="25"><div align="center"></div></td>
                  </tr>
                  <script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//--></script>
                <tr>
                  <td height="50"><div align="center">
                      <form name="links">
                        <select name="links" class="style2" onChange="window.open(this.value)">
                          <option value="javascript:void(null);" selected>----�����ѧ----</option>
                          <option value="http://www.harvard.edu/">�����ѧ</option>
                          <option value="http://www.cam.ac.uk/">���Ŵ�ѧ</option>
                          <option value="http://www.ox.ac.uk/">ţ���ѧ</option>
                          <option value="http://www.stanford.edu/">˹̹����ѧ</option>
                          <option value="http://www.yale.edu/">Ү³��ѧ</option>
                        </select>
                        </form>
                    </div></td>
                  </tr>
                  <tr>
                    <td height="50"><div align="center">
                        <form name="links">
                          <select name="links" class="style2" onChange="window.open(this.value)">
                            <option value="javascript:void(null);" selected>---ʵ��������---</option>
                            <option value="http://biophy.nju.edu.cn ">��������ʵ����</option>
                            <option value="http://pld.nju.edu.cn ">PLDʵ����</option>
                            <option value="http://x.nju.edu.cn/">�϶���С��</option>
                          </select>
                        </form>
                    </div></td>
                  </tr>
                  <tr>
                    <td height="50"><div align="center">
                        <form name="links">
                          <select name="links" class="style2" onChange="window.open(this.value)">
                            <option value="javascript:void(null);" selected>----У������----</option>
                            <option value="http://www.njbys.com/">�Ͼ���ҵ����ҵ��</option>
                            <option value="http://www.jsbys.com.cn/index.aspx">���ձ�ҵ����ҵ��</option>
                            <option value="http://www.firstjob.com.cn/">�Ϻ���ҵ����ҵ��</option>
                          </select>
                        </form>
                    </div></td>
                  </tr>
                </table>
            </div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="../links.asp" class="style3">&gt;&gt;&gt; More</a></div></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="34" background="../indeximages/loginbottom.gif">&nbsp;</td>
      </tr>
    </table></td>
    <td valign="top"><div align="right">
      <table width="603"  border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="406" height="10">&nbsp;</td>
            <td width="406" background="../indeximages/midLinkTop.gif">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td colspan="3"><div align="center">
              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="54" background="userimages/titlebk1.gif"><div align="center"><img src="userimages/stuinfo.gif" width="523" height="45"></div></td>
                  </tr>
                  <tr>
                    <td height="53" background="userimages/titlebk2.gif">&nbsp;</td>
                  </tr>
                  <tr>
                    <td background="userimages/titlebk.gif"><div align="center">
                      <form name="form" method="post" action="info_reg1.asp">
                        <div align="center">
                          <table width="90%" height="100%" border="0" cellpadding="0" cellspacing="0" class="thin">
                            <tr>
                              <td width="18%" height="24"><div align="right" class="style10">�û�����&nbsp;&nbsp;</div></td>
                              <td width="34%" class="style10"><input name="user_account" type="text" class="style3" value="<%=rs("user_account")%>" size="24">
                                  <span class="style13">*</span> </td>
                              <td width="48%" rowspan="9" valign="top" class="style10"><div align="center">
                                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td>
									    <div align="center">
									      <iframe src="denote.asp" name="denote" width="200" marginwidth="0" height="200" marginheight="0" align="middle" scrolling="yes" frameborder="1" allowtransparency="true" ></iframe>
          						      </div></td>
                                  </tr>
                                </table>
                              </div></td>
                            </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">���룺&nbsp;&nbsp;</div></td>
                              <td class="style10">
                                <input name="user_pwd" type="password" class="style3" size="24" >
                                  <span class="style13">*</span> </td>
                              </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">����ȷ�ϣ�&nbsp;&nbsp;</div></td>
                              <td class="style10"><input name="user_pwd1" type="password" class="style3"  size="24" >
                                  <span class="style13">*</span> </td>
                              </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">������&nbsp;&nbsp;</div></td>
                              <td class="style10"><%=rs("user_name")%>                                  </td>
                              </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">ѧ�ţ�&nbsp;&nbsp;</div></td>
                              <td class="style10"><%=rs("user_number")%></td>
                              </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">רҵ��&nbsp;&nbsp;</div></td>
                              <td class="style10"><%=rs("user_major")%></td>
                              </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">�꼶��&nbsp;&nbsp;</div></td>
                              <td class="style10">
                                  <select name="user_grade" class="style3" id="user_grade">
                                    <option value="��һ" selected>��һ</option>
                                    <option value="�ж�">�ж�</option>
                                    <option value="����">����</option>
									<option value="����">��һ</option>
									<option value="�ж�">����</option>
                                    <option value="����">����</option>
                                  </select>
                                  <span class="style13">*</span></td>
                              </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">E-mail��&nbsp;&nbsp;</div></td>
                              <td class="style10">
                                  <input name="user_mail" type="text" class="style3" value="<%=rs("user_mail")%>" size="24" >
                                  <span class="style13">*</span> </td>
                              </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">BBS�ʺţ�&nbsp;&nbsp;</div></td>
                              <td class="style10">
                                  <input name="user_bbs" type="text" class="style3" id="user_bbs" value="<%=rs("user_bbs")%>" size="24" >                                  </td>
                              </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">��ϵ��ʽ��&nbsp;&nbsp;</div></td>
                              <td colspan="2" class="style10">
                                  <input name="user_roomphone" type="text" class="style3" value="<%=rs("user_roomphone")%>" size="24" >
                                  <span class="style13">*</span></td>
                            </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">��ͥ�绰��&nbsp;&nbsp;</div></td>
                              <td colspan="2" class="style10">
                                  <input name="user_homephone" type="text" class="style3" value="<%=rs("user_homephone")%>" size="24" >
                                  <span class="style13">*</span>&nbsp;&nbsp;����ʽ��025-83594521�� </td>
                            </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">��ͥ��ַ��&nbsp;&nbsp;</div></td>
                              <td colspan="2" class="style10"><input name="user_address" type="text" class="style3" value="<%=rs("user_address")%>" size="24" >
                                <span class="style13">*</span>                              </td>
                            </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">�ʱࣺ&nbsp;&nbsp;</div></td>
                              <td colspan="2" class="style10">
                                  <input name="user_code" type="text" class="style3" value="<%=rs("user_code")%>" size="24" >
                                  <span class="style13">*</span>                              </td>
                            </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">�Ա�&nbsp;&nbsp;</div></td>
                              <td colspan="2" class="style10"><%=rs("user_sex")%></td>
                            </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">���գ�&nbsp;&nbsp;</div></td>
                              <td colspan="2" class="style10"><%=rs("user_Csrq")%></td>
                            </tr>
                            <tr>
                              <td height="24"><div align="right" class="style10">��ע��&nbsp;&nbsp;</div></td>
                              <td colspan="2" class="style10"><textarea name="user_info" cols="45" rows="10"><%=HTMLEncode(rs("user_info"))%></textarea>                              </td>
                            </tr>
                            <tr>
                              <td height="35" colspan="3"><div align="center"><img src="userimages/editSub.gif" width="70" height="25" align="absmiddle" style="cursor:hand; " onClick="javascript:checkuser()">
                              </div></td>
                            </tr>
                          </table>
                        </div>
                      </form>
                    </div></td>
                  </tr>
                  <tr>
                    <td height="34" background="userimages/titlebk3.gif">&nbsp;</td>
                  </tr>
                          </table>
            </div></td>
            </tr>
          </table>
    </div></td>
  </tr>
  <tr>
    <td height="15" valign="top">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top">
      <div align="right">
        <!--#include file="server.asp"-->        

      </div></td>
  </tr>
  <tr>
    <td height="12"></td>
    <td></td>
  </tr>
</table>
<!--#include file="bottom1.asp"-->
</html>
