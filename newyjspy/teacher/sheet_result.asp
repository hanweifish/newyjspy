<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("admin_account")="" then
Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
Response.end
end if
%>

<%
dim admin_account
admin_account=session("admin_account")
%>

<%
dim keywords,search_class
keywords=trim(request("keywords"))
search_class=trim(request("search_class"))
%>
<%
select case search_class
case "ѧ������"
search_class="user_info.user_name"
case "ѧ��ѧ��"
search_class="user_info.user_number"
case "�γ�"
search_class="sheet.course"
End select
%>
<%
set rs=server.createobject("adodb.recordset")
sql="select user_info.user_name,user_info.user_number,sheet.course,sheet.score,sheet.sheet_ID,sheet.property,sheet.sheet_info,subject.credit,subject.term from user_info inner join (sheet inner join subject on sheet.course = subject.course) on user_info.user_ID = sheet.user_ID where "&search_class&" like '%"&keywords&"%' order by subject.term,user_info.user_number"
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

<html>
<head>
<script language="javascript">
window.status="��ӭ�����Ͼ���ѧ����ϵ�о���������Ϣϵͳ��"
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
<!--#include file="top1.asp"-->
<table width="800"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="187" height="3"></td>
    <td rowspan="2" valign="top"><div align="right">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="5"><div align="right"></div></td>
        </tr>
        <tr>
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
                        <td height="54" background="adminimages/titlebk1.gif"><div align="center">
                            <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td><div align="center"><a href="sheet_set.asp"><img src="adminimages/scorequerry.gif" width="81" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="subject_set.asp"><img src="adminimages/examcourse.gif" width="81" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="sheet_add.asp"><img src="adminimages/scoreadd.gif" width="81" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="subject_add.asp"><img src="adminimages/courseadd1.gif" width="81" height="24" border="0"></a></div></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="86" background="adminimages/titlebk2.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="45"><div align="center"><img src="adminimages/scoremag.gif" width="523" height="45"></div></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr>
                        <td background="../user/userimages/titlebk.gif"><div align="center">
                            <table width="90%"  border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td><%
if (rs.bof and rs.eof) then
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>û���ҵ��κμ�¼!!!</font></marquee></td></tr>"
else
if search_class="user_info.user_name" or search_class="user_info.user_number"  then
%>
                                    <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                      <tr>
                                        <td height="25" colspan="10"><div align="center"><span class="style18">&quot;</span><span class="style14"><%=rs("user_name")%></span><span class="style18">&quot;ͬѧ�ɼ��� <span class="style19">��ѧ��</span></span> <span class="style5"><%=rs("user_number")%></span><span class="style18"><span class="style19">��</span></span></div></td>
                                      </tr>
                                      <tr>
                                        <td width="169" height="24"><div align="center"><span class="style10">�γ�</span></div></td>
                                        <td width="30" height="24"><div align="center" class="style10">�ɼ�</div></td>
                                        <td width="30" height="24"><div align="center" class="style10">ѧ��</div></td>
                                        <td width="60"><div align="center" class="style10">�γ�����</div></td>
                                        <td width="100"><div align="center" class="style10">�޶�ʱ��</div></td>
                                        <td width="80"><div align="center" class="style10">��ע��Ϣ</div></td>
                                        <td width="50"><div align="center" class="style10">�޸�</div></td>
                                        <td width="50" height="24"><div align="center" class="style10">ɾ��</div></td>
                                      </tr>
                                      <%
	for i=1 to rs.recordcount
%>
                                      <tr>
                                        <td height="24"><div align="center" class="style2">
                                            <div align="center" class="style2"><%=rs("course")%></div>
                                        </div></td>
                                        <td width="30" height="24"><div align="center" class="style10"><%=rs("score")%></div></td>
										<%
										dim thescore
										thescore=rs("score")
										if ISNumeric(rs("score"))  then
										
											if thescore>=60  then
											response.write "<td width=30 height=24><div align='center' class='style10'>"&rs("credit")&"</div></td>"
											else
											response.write "<td width=30 height=24><div align='center' class='style10'>0</div></td>"
											end if
										else 	
										response.write "<td width=30 height=24><div align='center' class='style10'>"&rs("credit")&"</div></td>"
										end if
										%>
                                        <td width="60"><div align="center" class="style10"><%=rs("property")%></div></td>
                                        <td width="100"><div align="center" class="style10"><%=rs("term")%></div></td>
                                        <td width="80"><div align="center" class="style10"><%=HTMLEncode(rs("sheet_info"))%></div></td>
                                        <td class="style3"><div align="center"><a href=sheet_mod.asp?sheet_ID=<%=rs("sheet_ID")%>>�޸�</a></div></td>
                                        <td width="50" height="24"class="style3"><div align="center"><a href=sheet_del.asp?sheet_ID=<%=rs("sheet_ID")%> >ɾ��</a></div></td>
                                      </tr>
                                      <%rs.movenext
next
rs.close
set rs=nothing
%>
                                    </table>
                                    <%
else
%>
                                    <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                      <tr>
                                        <td height="24" colspan="9"><div align="center"><span class="style18">&quot;</span><span class="style14"><%=rs("course")%></span><span class="style18">&quot;�γ̳ɼ���</span></div></td>
                                      </tr>
                                      <tr>
                                        <td width="50" height="24"><div align="center"><span class="style10">����</span></div></td>
                                        <td width="117"><div align="center"class="style10">ѧ�� </div></td>
                                        <td width="30" height="24"><div align="center" class="style10">�ɼ�</div></td>
                                        <td width="30" height="24"><div align="center" class="style10">ѧ��</div></td>
                                        <td width="60"><div align="center" class="style10">�γ�����</div></td>
                                        <td width="100"><div align="center" class="style10">�޶�ʱ��</div></td>
                                        <td width="80"><div align="center" class="style10">��ע��Ϣ</div></td>
                                        <td width="50"><div align="center" class="style10">�޸�</div></td>
                                        <td width="50" height="24"><div align="center" class="style10">ɾ��</div></td>
                                      </tr>
                                      <%
	for i=1 to rs.recordcount
%>
                                      <tr>
                                        <td width="50" height="24"><div align="center" class="style10">
                                            <div align="center" class="style10"><%=rs("user_name")%></div>
                                        </div></td>
                                        <td height="24"><div align="center"class="style10"><%=rs("user_number")%></div></td>
                                        <td width="30" height="24"><div align="center" class="style10"><%=rs("score")%></div></td>
                                        <td width="30" height="24"><div align="center" class="style10"><%=rs("credit")%></div></td>
                                        <td width="60"><div align="center" class="style10"><%=rs("property")%></div></td>
                                        <td width="100"><div align="center" class="style10"><%=HTMLEncode(rs("term"))%></div></td>
                                        <td width="80"><div align="center" class="style10"><%=rs("sheet_info")%></div></td>
                                        <td><div align="center" class="style18"><a href=sheet_mod.asp?sheet_ID=<%=rs("sheet_ID")%> class="style3">�޸�</a></div></td>
                                        <td width="50" height="24"><div align="center" class="style18"><a href=sheet_del.asp?sheet_ID=<%=rs("sheet_ID")%> class="style3">ɾ��</a></div></td>
                                      </tr>
                                      <%rs.movenext
next
rs.close
set rs=nothing
%>
                                    </table>
                                    <%
end if
end if
%>
                                </td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="40" background="adminimages/titlebk3.gif">&nbsp;</td>
                      </tr>
                    </table>
                </div></td>
              </tr>
            </table>
          </div></td>
        </tr>
        <tr>
          <td height="15" valign="top"><div align="right">
          </div></td>
        </tr>
        <tr>
          <td valign="top"><div align="right">
            <!--#include file="server.asp"-->
		  </div></td>
        </tr>
      </table>
    </div></td>
  </tr>
  <tr>
    <td valign="top" background="../indeximages/loginbk.gif"><div align="center">
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="47" background="adminimages/adminlogin.gif">&nbsp;</td>
          </tr>
          <tr>
            <td><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="100"><div align="center">
                    <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" background="../indeximages/loginbk.gif">
                      <tr>
                        <td width="20%"><div align="center"> </div>
                            <div align="center"></div></td>
                        <td width="60%" class="style3"><%=admin_account%>��</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr>
                        <td ><div align="center"> </div>
                            <div align="center"></div></td>
                        <td height="30" class="style2">���Ѿ���¼�ɹ�,��</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr>
                        <td><div align="center"></div></td>
                        <td height="30" class="style2">������ά����վ!</td>
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
                <td height="60" valign="center" background="../indeximages/loginbk.gif"><div align="center"><a href="admin_logout.asp"><img src="../includeimages/logout.gif" width="60" height="24" border="0"></a></div></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td height="77" background="../indeximages/links.gif">&nbsp;</td>
          </tr>
          <tr>
            <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
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
            <td>&nbsp;</td>
          </tr>
        </table>
      </div></td>
  </tr>
  <tr>
    <td height="34" background="../indeximages/loginbottom.gif">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<!--#include file="bottom1.asp"-->
</html>
