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
set rs=server.createobject("adodb.recordset")
sql="select * from teacher_info where admin_account='"&admin_account&"'"
rs.open sql,conn,1,1
admin_academy=rs("admin_academy")
%>
<script language="javascript">
	function checkform()
	{
		if (document.form.user_number.value=="")
		{
			alert("������ѧ�ţ�");
		}
		else if (document.form.train_to.value=="")
		{
			alert("������˳����䣡");
		}
	    else
		{
			return true;
		}
		return false;
	}
</script>
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
                                <td><div align="center"><img src="adminimages/train.gif" width="170" height="24"></div></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="86" background="adminimages/titlebk2.gif"><div align="center" class="style10">�˳�������д</div></td>
                      </tr>
                      <tr>
                        <td background="../user/userimages/titlebk.gif"><div align="center">
                            <table width="90%"  cellspacing="0" cellpadding="0">
                              <tr>
                                <td><div align="center">
                                    <table width="100%"  cellspacing="0" cellpadding="0">
                                      <tr>
                                        <td height="25"></td>
                                      </tr>
                                      <tr>
                                        <td><form action="train_infoadd1.asp" method="post" name="form" onSubmit="return checkform(form)">
                                            <div align="center">
                                              <table width="100%" border="1" cellpadding="0"  cellspacing="0" bordercolor="#000000" class="thin">
                                                <tr>
                                        <td height="24"><div align="right" class="style10">ѧ�ţ�&nbsp;&nbsp;</div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp;
                                            <%
									if session("user_number") <> "" then
									%>
                                            <input name="user_number" type="text" class="style3" value="<%=session("user_number")%>" >
                                            <%
									else
									%>
                                            <input name="user_number" type="text" class="style3" value="" >
                                            <%
									end if
									%>                                        </td>
                                      </tr>
										 <tr>
                                        <td height="24"><div align="right" class="style10">�˳�Ŀ�ĵأ�&nbsp;&nbsp;</div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp;
                                            <input name="train_to" type="text" class="style3" ></td>
                                      </tr>	
									  <tr>
                                        <td height="24"><div align="right" class="style10">������Ϣ��&nbsp;&nbsp;</div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp;
                                           <textarea name="train_info" cols="45" rows="10" class="style3">��</textarea></td>
                                      </tr>		
                                           <tr>
                                        <td height="24"><div align="right" class="style10">����Ժϵ��&nbsp;&nbsp;</div></td>
                                        <td class="style10">&nbsp;&nbsp;&nbsp;
                                            <%=rs("admin_yx")%><input type="hidden" name="train_academy" value="<%=rs("admin_academy")%>"></td>
                                      </tr>	     
                                                <tr>
                                                  <td height="25" colspan="2" class="style8 style12"><div align="center">
                                                      <input type="submit" name="Submit" value="ȷ ��">
&nbsp;&nbsp;&nbsp;</div></td>
                                                </tr>
                                              </table>
                                            </div>
                                        </form></td>
                                      </tr>
                                    </table>
                                </div></td>
                              </tr>
                              <tr>
                                <td><div align="center">
                                    <table width="100%"  border="0" cellpadding="0" cellspacing="0">
                                      <tr>
                                        <td height="25" colspan="6"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                          <%
								 set rsS = Server.CreateObject("adodb.recordset")
								 sqlS="select top 6 user_info.user_name,user_info.user_number,user_info.user_yx,user_info.user_pyxz,train_info.train_academy,train_info.train_info,train_info.train_ID,train_info.train_to from train_info inner join user_info on train_info.user_ID=user_info.user_ID where train_info.train_academy='"&admin_academy&"' order by train_info.train_id desc"
								 rsS.open sqlS,conn,1,1
								 %>
                                          <tr class="style10">
                                            <td height="22"><div align="center">����</div></td>
                                            <td><div align="center">ѧ��</div></td>
                                            <td><div align="center">�˳�Ŀ�ĵ�</div></td>
                                            <td><div align="center">Ժϵ</div></td>
											<td><div align="center">��������</div></td>
                                            <td><div align="center">������Ϣ</div></td>
                                            
                                            <td class="style3"><div align="center">�޸�</div></td>
                                            <td class="style3"><div align="center">ɾ��</div></td>
                                          </tr>
                                          <%
								 if not (rsS.eof and rsS.bof) then
								 for i=0 to rsS.recordcount
								 %>
                                          <tr class="style10">
                                            <td height="22"><div align="center"><%=rsS("user_name")%></div></td>
                                            <td><div align="center"><%=rsS("user_number")%></div></td>
                                            <td><div align="center"><%=rsS("train_to")%></div></td>
                                            <td><div align="center"><%=rsS("user_yx")%></div></td>
											<td><div align="center"><%=rsS("user_pyxz")%></div></td>
                                            <td><div align="center"><%=rsS("train_info")%></div></td>
                                           
                                            <td class="style3"><div align="center"><a href=train_infomod.asp?train_ID=<%=rsS("train_ID")%> class=style3 >�޸�</a></div></td>
                                            <td class="style3"><div align="center"><a href=train_infodel.asp?train_ID=<%=rsS("train_ID")%> class=style3 >ɾ��</a></div></td>
                                          </tr>
                                          <%
								  rsS.movenext
								  if rsS.eof then exit for 
								  next
								  end if
								  rsS.close
								  set rsS=nothing
								  %>
                                        </table></td>
                                      </tr>
                                      
                                      
                                     
                                    </table>
                                </div></td>
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
                        <td height="30" class="style2">����:<span class="style10"><%=rs("admin_yx")%></span></td>
                          </tr>
                          <tr>
						  <td><div align="center"></div></td>
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
<%
rs.close
set rs=nothing
%>
<!--#include file="bottom1.asp"-->
</html>
