<!--#include file="conn.asp"-->

<%
if request.cookies("status")="" then
    Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
	Response.end
end if
%>

<%
if session("admin_account")="" or session("user_group")="" then
Response.write"�Բ�������û�е�½���޴�Ȩ�ޣ�"
Response.end
end if
%>

<%
dim admin_account
admin_account=session("admin_account")
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

<%
dim course_id
course_id=trim(request("course_ID"))
session("course_ID")=course_id
set rs=server.createobject("adodb.recordset")
sql="select * from course where course_ID="&course_id
rs.open sql,conn,1,3
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
                                <td><div align="center"><a href="sel_search.asp"><img src="adminimages/querryresult.gif" width="134" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="course_set.asp"><img src="adminimages/selcourse.gif" width="134" height="24" border="0"></a></div></td>
                                <td><div align="center"><a href="course_add.asp"><img src="adminimages/courseadd.gif" width="134" height="24" border="0"></a></div></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="86" background="adminimages/titlebk2.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="45"><div align="center"><img src="adminimages/selset.gif" width="523" height="45"></div></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr>
                        <td background="../user/userimages/titlebk.gif"><div align="center">
                            <form name="form" method="post" action="course_mod1.asp">
                              <div align="center">
                                <table width="90%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
                                  <tr>
                                    <td height="24"><div align="right" class="style10">�γ����ƣ�&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input type="text" name="course_name" value="<%=rs("course_name")%>"></td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">�γ̱�ţ�&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input type="text" name="course_number" value="<%=rs("course_number")%>"></td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">�ڿ���ʦ��&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input type="text" name="course_tutor" value="<%=rs("course_tutor")%>"></td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">ѧ������&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input type="text" name="course_credit" value="<%=rs("course_credit")%>"></td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">�ڿ�ѧ�ڣ�&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <%
if rs("course_term") = "1" then
%>
                                        <select name="course_term">
                                          <option value="1" selected>ѡ�ογ�</option>
                                          <option value="0">��ѡ�ογ�</option>
                                        </select>
                                        <%
else
%>
                                        <select name="course_term">
                                          <option value="1">ѡ�ογ�</option>
                                          <option value="0" selected>��ѡ�ογ�</option>
                                        </select>
                                        <%
end if
%>
                                    </td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">�Ͽ�ʱ�䣺&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <input type="text" name="course_time" value="<%=rs("course_time")%>">
&nbsp;&nbsp;&nbsp;(��ʽ����һ5,6��)</td>
                                  </tr>
                                  <tr>
                                    <td height="24"><div align="right" class="style10">�γ̱�ע��&nbsp;&nbsp;</div></td>
                                    <td class="style10">&nbsp;&nbsp;&nbsp;
                                        <textarea name="course_info" cols="45" rows="10"><%=HTMLEncode(rs("course_info"))%></textarea>
                                    </td>
                                  </tr>
                                  <tr>
                                    <td height="35" colspan="2"><div align="center">
                                        <input type="submit" name="Submit" value="�ύ����">
                                    </div></td>
                                  </tr>
                                </table>
                              </div>
                            </form>
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
