<!--#include file="conn.asp"-->

<%
set rs=server.createobject("adodb.recordset")
sql="select * from tutor order by tutor_major , tutor_name"
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
window.status="��ӭ�����Ͼ���ѧ����ϵ�о���������Ϣϵͳ��"
</script>
<script language="javascript">
	function checkuser(form)
	{
		if (document.form.user_account.value=="")
		{
			alert("�������û�������");
		}
		else if (document.form.user_pwd.value=="")
		{
			alert("���������룡!");
		}
		else
		{
			form.submit();
		}
		return false;
	}
</script>

<script language="javascript">
	function checkuser1()
	{
		if (document.form1.user_account.value=="")
		{
			alert("�������û�����");
		}
		else if (document.form1.user_pwd.value=="")
		{
			alert("���������룡");
		}
		else if (document.form1.user_name.value=="")
		{
			alert("������������");
		}
		else if (document.form1.user_number.value=="")
		{
			alert("������ѧ�ţ�");
		}
		else if (document.form1.user_mail.value=="")
		{
			alert("������������䣡");
		}
		else if (document.form1.user_roomphone.value=="")
		{
			alert("����������绰��");
		}
		else
		{
			return true;
		}
		return false;
	}
</script>
<script language="javascript">
	function checknumber()
	{
	var user_number=document.form1.user_number.value,yanyi_bg=document.form1.yanyi_bg.value,yanyi_end=document.form1.yanyi_end.value,yaner_bg=document.form1.yaner_bg.value,yaner_end=document.form1.yaner_end.value,yansan_bg=document.form1.yansan_bg.value,yansan_end=document.form1.yansan_end.value
	if(user_number.length!=9 || !((user_number>=yanyi_bg&&user_number<=yanyi_end)||(user_number>=yaner_bg&&user_number<=yaner_end)||(user_number>=yansan_bg&&user_number<=yansan_end))) {
		alert("�Բ��������߱�ע��Ȩ�ޣ�")
		return false
		}
	return true
	}
</script>
<script language="javascript">
function submitform(form1){
	if(checkuser1()&&checknumber())
		form1.submit();
	else
		return false;
}
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�о�����Ϣ����ϵͳ</title>
<link href="style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style13 {font-size: 24px}
-->
</style>
<!--#include file="top.asp"-->
<table width="800"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="187" height="12"></td>
    <td></td>
  </tr>
  <tr>
    <td rowspan="2" valign="top"><table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="47" background="indeximages/rcpy.gif">&nbsp;</td>
      </tr>
      <tr>
        <td><div align="center">
          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="350" valign="middle" background="indeximages/loginbk.gif"><div align="center">
                  <table width="70%" height="80%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td scope="col"><div align="center"><a href="tutor.asp" class="style2">��ʦ���</a></div></td>
                    </tr>
                    <tr>
                      <td><div align="center"><a href="yjspy.asp" class="style3">�о�������</a></div></td>
                    </tr>
                    <tr>
                      <td><div align="center"><a href="yjsxw.asp" class="style3">�о���ѧλ</a></div></td>
                    </tr>
                    <tr>
                      <td><div align="center"><a href="yjsxj.asp" class="style3">�о���ѧ��</a></div></td>
                    </tr>
                    <tr>
                      <td><div align="center"><a href="yjsjx.asp" class="style3">�о�����ѧ</a></div></td>
                    </tr>
                  </table>
              </div></td>
            </tr>
          </table>
        </div></td>
      </tr>
      <tr>
        <td height="77" background="indeximages/links.gif">&nbsp;</td>
      </tr>
      <tr>
        <td background="indeximages/loginbk.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
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
            <td height="35"><div align="center"><a href="links.asp" class="style3">&gt;&gt;&gt; More</a></div></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="34" background="indeximages/loginbottom.gif">&nbsp;</td>
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
                  <td height="54" background="user/userimages/titlebk1.gif"><div align="center"><img src="indeximages/tutor.gif" width="523" height="45"></div></td>
                </tr>
                <tr>
                  <td height="53" valign="bottom" background="user/userimages/titlebk2.gif"><div align="center">�����������Ⱥ󣬷�רҵ�����������У�</div></td>
                </tr>
                <tr>
                  <td background="user/userimages/titlebk.gif"><div align="center" class="style13">
                    <table width="90%"  border="0" cellspacing="0" cellpadding="0">
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
			NoncePage=request("page")
		end if
		if Cint(Trim(request("page")))>Cint(NumPage) then NoncePage=NumPage
	end if
else
	NumRecord=0
	NumPage=0
	NoncePage=0
	end if
%>
                      <tr>
                        <td><div align="center">
                            <table width="100%" height="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000"class="thin">
                              <tr>
                                <td width="110" height="25" bgcolor="#DDE6F4"><div align="center" class="style2">
                                  <div align="center">��ʦ����</div>
                                </div></td>
                                <td width="245"><div align="center" class="style2"><div align="center">
                                  <div align="center" class="style2">
                                    <div align="center"></div>
                                  </div>
                                  רҵ����</div>
                                </div></td>
                                	<td width="85"><div align="center" class="style2">ְ&nbsp;&nbsp; �� </div></td>
                                	<td width="93"><div align="center"><span class="style2">����/˶��</span></div></td>
                              </tr>
<%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*20,1
	for i=1 to rs.pagesize
%>
							  <tr>
                                <td height="25" bgcolor="#DDE6F4"><div align="center">
                                  <div align="center" class="style2">
                                    <div align="center"><a href="tutor_detail.asp?tutor_id=<%=rs("tutor_id")%>"><%=rs("tutor_name")%></a></div>
                                  </div>
                                  </div></td>
                                <td><div align="center" class="style2"><%=rs("tutor_major")%></div></td>
                                <td><div align="center">
                                  <div align="center" class="style2">
                                    <div align="center"><%=rs("tutor_post")%></div>
                                  </div>
                                </div></td>
                                <td><div align="center">
                                  <div align="center"><span class="style2"><%=rs("tutor_mord")%></span></div>
                                </div></td>
                              </tr>
                      <%rs.movenext
if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13 height='24'><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>��ʱû�м�¼!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
                            </table>
                        </div></td>
                      </tr>
                      <tr>
                        <td height="20" colspan="8" valign="middle"><div align="right"> <span class="style2">
                            <input type="hidden" name="page" value="<%=NoncePage%>">
                            <%
if NoncePage>1 then
	response.write "|<a href=tutor.asp?page=1>�� ҳ</a>| |<a href=tutor.asp?page="&NoncePage-1&">��һҳ</a>|&nbsp"
else
	response.write "|�� ҳ| |��һҳ|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=tutor.asp?page="&NoncePage+1&">��һҳ</a>| |<a href=tutor.asp?page="&NumPage&">β ҳ</a>|"
else
	response.write "|��һҳ| |β ҳ|"
end if
%>
&nbsp;ҳ�Σ�<font color="#0033CC"><%=NoncePage%></font>/<font color="#0033CC"><%=NumPage%></font> ��<font color="#0033CC"><%=NumRecord%></font>����¼</span>&nbsp; </div></td>
                      </tr>
                    </table>
                    </div></td>
                </tr>
                <tr>
                  <td height="34" background="user/userimages/titlebk3.gif">&nbsp;</td>
                </tr>
              </table>
          </div></td>
        </tr>
      </table>
    </div></td>
  </tr>
  <tr>
    <td rowspan="2" valign="top">&nbsp;</td>
  </tr>
  <tr>
    <td height="12"></td>
    </tr>
</table>
<!--#include file="bottom.asp"-->
</html>
