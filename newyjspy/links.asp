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
                  <td height="54" background="user/userimages/titlebk1.gif"><div align="center"><img src="indeximages/linkstitle.gif" width="523" height="45"></div></td>
                </tr>
                <tr>
                  <td height="53" background="user/userimages/titlebk2.gif">&nbsp;</td>
                </tr>
                <tr>
                  <td height="450" valign="top" background="user/userimages/titlebk.gif"><div align="center">
                    <table width="90%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="55" scope="col"><div align="center">
                          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="15%" rowspan="2" bgcolor="#F3F7FC"  scope="col"><div align="center" class="style3">У������</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://www.nju.edu.cn/">�ϴ���ҳ</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://grawww.nju.edu.cn/">�о���Ժ</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://lib.nju.edu.cn/">ͼ���</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://info.nju.edu.cn/">���ֻ�У԰</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><SPAN class=main><FONT color=#336699><a href="http://www.nju.edu.cn/cps/site/NJU/njuc/lx.htm">Ժϵ���緽��</a></FONT></SPAN></td>
                            </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><a href="http://tuanwei.nju.edu.cn/njuyh/">�о�����</a></td>
                              <td bgcolor="#E6ECF7"><a href="http://www.nju.edu.cn/cps/site/NJU/njuc/timetable/index.htm">�г�ʱ��</a></td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                            </tr>
                          </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td height="55"><div align="center">
                          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="15%" rowspan="2" bgcolor="#E6ECF7"  scope="col"><div align="center" class="style3">��ҵ���</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://job.nju.edu.cn/">�ϴ��ҵָ������</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://www.njbys.com/">�Ͼ���ҵ����ҵ��</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://www.jsbys.com.cn/index.aspx">���ձ�ҵ����ҵ��</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://www.js.lm.gov.cn/">�й��Ͷ����г���</a></td>
                              </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><a href="http://www.firstjob.com.cn/">�Ϻ���ҵ����ҵ��</a></td>
                              <td bgcolor="#E6ECF7"><a href="http://www.chinahr.com/">�л�Ӣ����</a></td>
                              <td bgcolor="#E6ECF7"><a href="http://www.js.lm.gov.cn/">����ʡְҵ��������</a></td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                              </tr>
                          </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td height="55"><div align="center">
                          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="15%" rowspan="2" bgcolor="#F3F7FC"  scope="col"><div align="center" class="style3">��������</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://www.google.com/">www.google.com</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://www.baidu.com/">�ٶ�����</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://thephy.nju.edu.cn/">ThePhy����</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://e.pku.edu.cn/">������������</a></td>
                            </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><a href="http://www.sheenk.com/">�ǿ�����</a></td>
                              <td bgcolor="#E6ECF7"><a href="http://sesa.nju.edu.cn/search/">����ϵʯͷ��</a></td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                            </tr>
                          </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td height="55"><div align="center">
                          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="15%" rowspan="2"  scope="col"><div align="center" class="style3">��������</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://news.sina.com.cn/">��������</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://news.sohu.com/">�Ѻ�����</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A href="http://www.xinhuanet.com/">�»���</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://www.peopledaily.com.cn/">�����ձ�</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://www.chinanews.com.cn/">������</A></td>
                            </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><A 
                        href="http://news.tom.com/">TOM����</A></td>
                              <td bgcolor="#E6ECF7"><a href="http://www.nanfangdaily.com.cn/zm/20050106/">�Ϸ���ĩ</a></td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                            </tr>
                          </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td height="55"><div align="center">
                          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="15%" rowspan="2" bgcolor="#F3F7FC"  scope="col"><div align="center" class="style3">�������</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://www.skycn.com/">������</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://www.onlinedown.net/">�������԰</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://download.21cn.com/">21cn����</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://download.pchome.net/">����֮��</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A href="http://www.pconline.com.cn/download/" 
                        target=_blank>̫ƽ������</A></td>
                            </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><A 
                        href="http://download.sina.com.cn/">��������</A></td>
                              <td bgcolor="#E6ECF7"><a href="http://dl.163.com/">�����������</a></td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                            </tr>
                          </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td height="55"><div align="center">
                          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="15%" rowspan="2"  scope="col"><div align="center" class="style3">����Mp3</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
      href="http://www.mtvtop.com/">�й���������</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
      href="http://game.baidu.com/traf/click.php?id=579&amp;url=http://www.9sky.com/">��������</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
      onmouseover="window.status='';return true" 
      href="http://www.hao123.com/fra/music/tyfocom.htm">�컢������</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
      href="http://www.hao123.com/music/chinayy.htm" target=_blank>�й�Ӱ������</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
      href="http://www.99music.net/new.htm">�þ�����</A></td>
                            </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><A 
      href="http://www.hao123.com/fra/music/chinamp3.htm">���ּ���</A></td>
                              <td bgcolor="#E6ECF7"><A 
      href="http://www.hao123.com/fra/music/soguacom.htm" 
      target=_blank>MP3�ѹ���</A></td>
                              <td bgcolor="#E6ECF7"><A 
      href="http://music.feifa.com/">�������ֹ�</A></td>
                              <td bgcolor="#E6ECF7"><A 
      href="http://www.etang.com/music/">��������</A></td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                            </tr>
                          </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td height="55"><div align="center">
                          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="15%" rowspan="2" bgcolor="#F3F7FC"  scope="col"><div align="center" class="style3">����Ƶ��</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://sports.sina.com.cn/">���˾����籩</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://sports.163.com/">��������Ƶ��</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://sports.sohu.com/" 
                        target=_blank>�Ѻ�����</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://www.cctv.com/sports/">cctv����Ƶ��</A></td>
                              </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><A 
                        href="http://sports.tom.com/">TOM������̳</A></td>
                              <td bgcolor="#E6ECF7"><A 
      href="http://www.beijing-olympic.org.cn/" 
      target=_blank>2008���˹ٷ�</A></td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                              </tr>
                          </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td bgcolor="#F3F7FC"><div align="center"><a href="university.asp" class="style2">| ������ֱ����Уһ�� |</a>
                          </div></td>
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
