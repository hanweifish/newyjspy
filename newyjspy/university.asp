<!--#include file="conn.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select top 8 * from notice order by notice_time desc"
rs.open sql,conn,1,1
%>
<%
set rs=createobject("adodb.recordset")
sql="select * from ynumber"
rs.open sql,conn,1,1
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
.style10 {font-size: 12px;
	color: #004080;
}
-->
</style>
<style type="text/css">
<!--
.style12 {color: #FF0000}
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
        <td height="47" background="indeximages/stulogin.gif">&nbsp;</td>
      </tr>
      <tr>
        <td><div align="center">
          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="132"><div align="center">
                <form action="user/user_check.asp" method="post" name="form" id="form">
                  <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" background="indeximages/loginbk.gif">
                    <tr>
                      <td colspan="2"><div align="center">�û���:
                          <input name="user_account" type="text" class="style3" id="name" size="12">
                      </div>
                        <div align="left">                        </div></td>
                      </tr>
                    <tr>
                      <td height="38" colspan="2"><div align="center">�� &nbsp;��:
                          <input name="user_pwd" type="password" class="style3" id="pwd" size="12"> 
                        </div>
                        <div align="left">                        </div></td>
                      </tr>
                    <tr>
                      <td width="50%"><div align="right"><img src="indeximages/login.gif" width="49" height="23" border="0" style='cursor:hand' onMouseDown="checkuser(form)">&nbsp;</div></td>
                      <td width="50%"><div align="left">&nbsp;<a href="javascript:void(null)"><img src="indeximages/register.gif" width="49" height="23" border="0"></a></div></td>
                    </tr>
                    <tr>
                      <td height="10"><div align="center"></div></td>
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                </form>
                </div></td>
            </tr>
            <tr>
              <td height="6" background="indeximages/loginbk.gif"><div align="center"><img src="indeximages/loginbar.gif" width="129" height="2"></div></td>
            </tr>
            <tr>
              <td height="120" valign="center" background="indeximages/loginbk.gif"><div align="center"><iframe src="denote.asp" name="denote" width="150" marginwidth="0" height="120" marginheight="0" align="middle" scrolling="yes" frameborder="0" allowtransparency="true" ></iframe>
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
                  <td height="54" background="user/userimages/titlebk1.gif"><div align="center"><img src="indeximages/university.gif" width="523" height="45"></div></td>
                </tr>
                <tr>
                  <td height="53" background="user/userimages/titlebk2.gif">&nbsp;</td>
                </tr>
                <tr>
                  <td height="450" valign="top" background="user/userimages/titlebk.gif"><div align="center">
                    <TABLE width="90%" height="100%" 
border=0 align=center cellPadding=0 cellSpacing=0 class="thin">
                      <TBODY>
                        <TR>
                          <TD height="25"><A href="http://www.tsinghua.edu.cn/" 
                  target=_blank>�廪��ѧ</A></TD>
                          <TD><A href="http://www.pku.edu.cn/" target=_blank>������ѧ</A></TD>
                          <TD><A href="http://www.ruc.edu.cn/" 
                target=_blank>�й������ѧ</A></TD>
                          <TD><A href="http://www.ustb.edu.cn/" 
                target=_blank>�����Ƽ���ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.buct.edu.cn/" 
                target=_blank>����������ѧ</A></TD>
                          <TD><A href="http://www.bnu.edu.cn/" 
                target=_blank>����ʦ����ѧ</A></TD>
                          <TD><A href="http://www.bfsu.edu.cn/" 
                target=_blank>����������ѧ</A></TD>
                          <TD><A href="http://www.blcu.edu.cn/" 
                  target=_blank>���������Ļ���ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.nankai.edu.cn/" 
                target=_blank>�Ͽ���ѧ</A></TD>
                          <TD><A href="http://www.tju.edu.cn/" target=_blank>����ѧ</A></TD>
                          <TD><A href="http://www.neu.edu.cn/" target=_blank>������ѧ</A></TD>
                          <TD><A href="http://www.dlut.edu.cn/" 
                target=_blank>��������ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.jlu.edu.cn/" target=_blank>���ִ�ѧ</A></TD>
                          <TD><A href="http://www.znufe.edu.cn/" 
                  target=_blank>���ϲƾ�������ѧ</A></TD>
                          <TD><A href="http://www.nenu.edu.cn/" 
                target=_blank>����ʦ����ѧ</A></TD>
                          <TD><A href="http://www.fudan.edu.cn/" 
                  target="_bla??&#4;nk">������ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.sjtu.edu.cn/" 
                target=_blank>�Ϻ���ͨ��ѧ</A></TD>
                          <TD><A href="http://www.tongji.edu.cn/" 
                target=_blank>ͬ�ô�ѧ</A></TD>
                          <TD><A href="http://www.ecust.edu.cn/" 
                target=_blank>��������ѧ</A></TD>
                          <TD><A href="http://www.xidian.edu.cn/" 
                  target=_blank>�������ӿƼ���ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.ecnu.edu.cn/" 
                target=_blank>����ʦ����ѧ</A></TD>
                          <TD><A href="http://www.shisu.edu.cn/" 
                  target=_blank>�Ϻ�������ѧ</A></TD>
                          <TD><A href="http://www.nju.edu.cn/" target=_blank class="style3">�Ͼ���ѧ</A></TD>
                          <TD><A href="http://www.seu.edu.cn/" 
              target=_blank>���ϴ�ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.wxuli.edu.cn/" 
                target=_blank>�����Ṥ��ѧ</A></TD>
                          <TD><A href="http://www.hfut.edu.cn/" 
                target=_blank>�Ϸʹ�ҵ��ѧ</A></TD>
                          <TD><A href="http://www.zju.edu.cn/" target=_blank>�㽭��ѧ</A></TD>
                          <TD><A href="http://www.xmu.edu.cn/" 
              target=_blank>���Ŵ�ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.sdu.edu.cn/" target=_blank>ɽ����ѧ</A></TD>
                          <TD><A href="http://www.ouqd.edu.cn/" 
                target=_blank>�ൺ�����ѧ</A></TD>
                          <TD><A href="http://www.whu.edu.cn/" target=_blank>�人��ѧ</A></TD>
                          <TD><A href="http://www.swjtu.edu.cn/" 
                target=_blank>���Ͻ�ͨ��ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.cug.edu.cn/" 
                target=_blank>�й����ʴ�ѧ</A></TD>
                          <TD><A href="http://www.ccnu.edu.cn/" 
                target=_blank>����ʦ����ѧ</A></TD>
                          <TD><A href="http://www.hunu.edu.cn/" 
target=_blank>���ϴ�ѧ</A></TD>
                          <TD><A href="http://www.cpu.edu.cn/" 
                target=_blank>�й�ҩ�ƴ�ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.zsu.edu.cn/" target=_blank>��ɽ��ѧ</A></TD>
                          <TD><A href="http://www.scut.edu.cn/" 
                target=_blank>��������ѧ</A></TD>
                          <TD><A href="http://www.lzu.edu.cn/" target=_blank>���ݴ�ѧ</A></TD>
                          <TD><A href="http://www.uestc.edu.cn/" 
                target=_blank>���ӿƼ���ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.scu.edu.cn/" target=_blank>�Ĵ���ѧ</A></TD>
                          <TD><A href="http://www.cqu.edu.cn/" target=_blank>�����ѧ</A></TD>
                          <TD><A href="http://www.swnu.edu.cn/" 
                target=_blank>����ʦ����ѧ</A></TD>
                          <TD><A href="http://www.xjtu.edu.cn/" 
                target=_blank>������ͨ��ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.snnu.edu.cn/" 
                target=_blank>����ʦ����ѧ</A></TD>
                          <TD><A href="http://www.swufe.edu.cn/" 
                target=_blank>���ϲƾ���ѧ</A></TD>
                          <TD><A href="http://www.nwsuaf.edu.cn/" 
                  target=_blank>����ũ�ֿƼ���ѧ</A></TD>
                          <TD><A href="http://www.uibe.edu.cn/" 
                  target=_blank>���⾭��ó�״�ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.bjpeu.edu.cn/" 
                target=_blank>ʯ�ʹ�ѧ</A></TD>
                          <TD><A href="http://www.bupt.edu.cn/" 
                target=_blank>�����ʵ��ѧ</A></TD>
                          <TD><A href="http://www.cau.edu.cn/" 
                target=_blank>�й�ũҵ��ѧ</A></TD>
                          <TD><A href="http://www.bjfu.edu.cn/" 
                target=_blank>������ҵ��ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.bbi.edu.cn/" 
                target=_blank>�����㲥ѧԺ</A></TD>
                          <TD><A href="http://www.hhu.edu.cn/" target=_blank>�Ӻ���ѧ</A></TD>
                          <TD><A href="http://www.cupl.edu.cn/" 
                target=_blank>�й�������ѧ</A></TD>
                          <TD><A href="http://www.ccom.edu.cn/" 
                target=_blank>��������ѧԺ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.cumt.edu.cn/" 
                target=_blank>�й���ҵ��ѧ</A></TD>
                          <TD><A href="http://www.shufe.edu.cn/" 
                target=_blank>�Ϻ��ƾ���ѧ</A></TD>
                          <TD><A href="http://www.nefu.edu.cn/" 
                target=_blank>������ҵ��ѧ</A></TD>
                          <TD><A href="http://www.hzau.edu.cn/" 
                target=_blank>����ũҵ��ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.njau.edu.cn/" 
                target=_blank>�Ͼ�ũҵ��ѧ</A></TD>
                          <TD><A href="http://www.dhu.edu.cn/" target=_blank>������ѧ</A></TD>
                          <TD><A href="http://www.chntheatre.edu.cn/" 
                  target=_blank>����Ϸ��ѧԺ</A></TD>
                          <TD><A href="http://www.xahu.edu.cn/" 
              target=_blank>������ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.njtu.edu./??&#4;cn" 
                  target=_blank>������ͨ��ѧ</A></TD>
                          <TD><A href="http://www.hust.edu.cn/" 
                target=_blank>���пƼ���ѧ</A></TD>
                          <TD><A href="http://www.csu.edu.cn/" target=_blank>���ϴ�ѧ</A></TD>
                          <TD><A href="http://www.cufe.edu.cn/" 
                target=_blank>����ƾ���ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.whut.edu.cn/" 
                target=_blank>�人����ѧ</A></TD>
                          <TD><A href="http://www.china-visual.com/meiyuan/"
				target=_blank>��������ѧԺ</A></TD>
                          <TD><A href="http://www.bjucmp.edu.cn/" 
                  target=_blank>������ҽҩ��ѧ</A></TD>
                          <TD><A href="http://www.ncepu.edu.cn/indexp.htm" 
                  target=_blank>����������ѧ</A></TD>
                        </TR>
                        <TR>
                          <TD height="30" colspan="4"><div align="center"><a href="links.asp"><img src="user/userimages/return.gif" width="49" height="23" border="0" align="absmiddle"></a></div></TD>
                          </TR>
                      </TBODY>
                    </TABLE>
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
