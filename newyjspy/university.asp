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
window.status="欢迎访问南京大学物理系研究生管理信息系统！"
</script>
<script language="javascript">
	function checkuser(form)
	{
		if (document.form.user_account.value=="")
		{
			alert("请输入用户名！！");
		}
		else if (document.form.user_pwd.value=="")
		{
			alert("请输入密码！!");
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
			alert("请输入用户名！");
		}
		else if (document.form1.user_pwd.value=="")
		{
			alert("请输入密码！");
		}
		else if (document.form1.user_name.value=="")
		{
			alert("请输入姓名！");
		}
		else if (document.form1.user_number.value=="")
		{
			alert("请输入学号！");
		}
		else if (document.form1.user_mail.value=="")
		{
			alert("请输入电子邮箱！");
		}
		else if (document.form1.user_roomphone.value=="")
		{
			alert("请输入宿舍电话！");
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
		alert("对不起，您不具备注册权限！")
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
<title>研究生信息管理系统</title>
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
                      <td colspan="2"><div align="center">用户名:
                          <input name="user_account" type="text" class="style3" id="name" size="12">
                      </div>
                        <div align="left">                        </div></td>
                      </tr>
                    <tr>
                      <td height="38" colspan="2"><div align="center">密 &nbsp;码:
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
                          <option value="javascript:void(null);" selected>----国外大学----</option>
                          <option value="http://www.harvard.edu/">哈佛大学</option>
                          <option value="http://www.cam.ac.uk/">剑桥大学</option>
                          <option value="http://www.ox.ac.uk/">牛津大学</option>
                          <option value="http://www.stanford.edu/">斯坦福大学</option>
                          <option value="http://www.yale.edu/">耶鲁大学</option>
                        </select>
                        </form>
                    </div></td>
                  </tr>
                  <tr>
                    <td height="50"><div align="center">
                        <form name="links">
                          <select name="links" class="style2" onChange="window.open(this.value)">
                            <option value="javascript:void(null);" selected>---实验室链接---</option>
                            <option value="http://biophy.nju.edu.cn ">生物物理实验室</option>
                            <option value="http://pld.nju.edu.cn ">PLD实验室</option>
                            <option value="http://x.nju.edu.cn/">邢定钰小组</option>
                          </select>
                        </form>
                    </div></td>
                  </tr>
                  <tr>
                    <td height="50"><div align="center">
                        <form name="links">
                          <select name="links" class="style2" onChange="window.open(this.value)">
                            <option value="javascript:void(null);" selected>----校外链接----</option>
                            <option value="http://www.njbys.com/">南京毕业生就业网</option>
                            <option value="http://www.jsbys.com.cn/index.aspx">江苏毕业生就业网</option>
                            <option value="http://www.firstjob.com.cn/">上海毕业生就业网</option>
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
                  target=_blank>清华大学</A></TD>
                          <TD><A href="http://www.pku.edu.cn/" target=_blank>北京大学</A></TD>
                          <TD><A href="http://www.ruc.edu.cn/" 
                target=_blank>中国人民大学</A></TD>
                          <TD><A href="http://www.ustb.edu.cn/" 
                target=_blank>北京科技大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.buct.edu.cn/" 
                target=_blank>北京化工大学</A></TD>
                          <TD><A href="http://www.bnu.edu.cn/" 
                target=_blank>北京师范大学</A></TD>
                          <TD><A href="http://www.bfsu.edu.cn/" 
                target=_blank>北京外国语大学</A></TD>
                          <TD><A href="http://www.blcu.edu.cn/" 
                  target=_blank>北京语言文化大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.nankai.edu.cn/" 
                target=_blank>南开大学</A></TD>
                          <TD><A href="http://www.tju.edu.cn/" target=_blank>天津大学</A></TD>
                          <TD><A href="http://www.neu.edu.cn/" target=_blank>东北大学</A></TD>
                          <TD><A href="http://www.dlut.edu.cn/" 
                target=_blank>大连理工大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.jlu.edu.cn/" target=_blank>吉林大学</A></TD>
                          <TD><A href="http://www.znufe.edu.cn/" 
                  target=_blank>中南财经政法大学</A></TD>
                          <TD><A href="http://www.nenu.edu.cn/" 
                target=_blank>东北师范大学</A></TD>
                          <TD><A href="http://www.fudan.edu.cn/" 
                  target="_bla??&#4;nk">复旦大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.sjtu.edu.cn/" 
                target=_blank>上海交通大学</A></TD>
                          <TD><A href="http://www.tongji.edu.cn/" 
                target=_blank>同济大学</A></TD>
                          <TD><A href="http://www.ecust.edu.cn/" 
                target=_blank>华东理工大学</A></TD>
                          <TD><A href="http://www.xidian.edu.cn/" 
                  target=_blank>西安电子科技大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.ecnu.edu.cn/" 
                target=_blank>华东师范大学</A></TD>
                          <TD><A href="http://www.shisu.edu.cn/" 
                  target=_blank>上海外国语大学</A></TD>
                          <TD><A href="http://www.nju.edu.cn/" target=_blank class="style3">南京大学</A></TD>
                          <TD><A href="http://www.seu.edu.cn/" 
              target=_blank>东南大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.wxuli.edu.cn/" 
                target=_blank>无锡轻工大学</A></TD>
                          <TD><A href="http://www.hfut.edu.cn/" 
                target=_blank>合肥工业大学</A></TD>
                          <TD><A href="http://www.zju.edu.cn/" target=_blank>浙江大学</A></TD>
                          <TD><A href="http://www.xmu.edu.cn/" 
              target=_blank>厦门大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.sdu.edu.cn/" target=_blank>山东大学</A></TD>
                          <TD><A href="http://www.ouqd.edu.cn/" 
                target=_blank>青岛海洋大学</A></TD>
                          <TD><A href="http://www.whu.edu.cn/" target=_blank>武汉大学</A></TD>
                          <TD><A href="http://www.swjtu.edu.cn/" 
                target=_blank>西南交通大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.cug.edu.cn/" 
                target=_blank>中国地质大学</A></TD>
                          <TD><A href="http://www.ccnu.edu.cn/" 
                target=_blank>华中师范大学</A></TD>
                          <TD><A href="http://www.hunu.edu.cn/" 
target=_blank>湖南大学</A></TD>
                          <TD><A href="http://www.cpu.edu.cn/" 
                target=_blank>中国药科大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.zsu.edu.cn/" target=_blank>中山大学</A></TD>
                          <TD><A href="http://www.scut.edu.cn/" 
                target=_blank>华南理工大学</A></TD>
                          <TD><A href="http://www.lzu.edu.cn/" target=_blank>兰州大学</A></TD>
                          <TD><A href="http://www.uestc.edu.cn/" 
                target=_blank>电子科技大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.scu.edu.cn/" target=_blank>四川大学</A></TD>
                          <TD><A href="http://www.cqu.edu.cn/" target=_blank>重庆大学</A></TD>
                          <TD><A href="http://www.swnu.edu.cn/" 
                target=_blank>西南师范大学</A></TD>
                          <TD><A href="http://www.xjtu.edu.cn/" 
                target=_blank>西安交通大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.snnu.edu.cn/" 
                target=_blank>陕西师范大学</A></TD>
                          <TD><A href="http://www.swufe.edu.cn/" 
                target=_blank>西南财经大学</A></TD>
                          <TD><A href="http://www.nwsuaf.edu.cn/" 
                  target=_blank>西北农林科技大学</A></TD>
                          <TD><A href="http://www.uibe.edu.cn/" 
                  target=_blank>对外经济贸易大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.bjpeu.edu.cn/" 
                target=_blank>石油大学</A></TD>
                          <TD><A href="http://www.bupt.edu.cn/" 
                target=_blank>北京邮电大学</A></TD>
                          <TD><A href="http://www.cau.edu.cn/" 
                target=_blank>中国农业大学</A></TD>
                          <TD><A href="http://www.bjfu.edu.cn/" 
                target=_blank>北京林业大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.bbi.edu.cn/" 
                target=_blank>北京广播学院</A></TD>
                          <TD><A href="http://www.hhu.edu.cn/" target=_blank>河海大学</A></TD>
                          <TD><A href="http://www.cupl.edu.cn/" 
                target=_blank>中国政法大学</A></TD>
                          <TD><A href="http://www.ccom.edu.cn/" 
                target=_blank>中央音乐学院</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.cumt.edu.cn/" 
                target=_blank>中国矿业大学</A></TD>
                          <TD><A href="http://www.shufe.edu.cn/" 
                target=_blank>上海财经大学</A></TD>
                          <TD><A href="http://www.nefu.edu.cn/" 
                target=_blank>东北林业大学</A></TD>
                          <TD><A href="http://www.hzau.edu.cn/" 
                target=_blank>华中农业大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.njau.edu.cn/" 
                target=_blank>南京农业大学</A></TD>
                          <TD><A href="http://www.dhu.edu.cn/" target=_blank>东华大学</A></TD>
                          <TD><A href="http://www.chntheatre.edu.cn/" 
                  target=_blank>中央戏剧学院</A></TD>
                          <TD><A href="http://www.xahu.edu.cn/" 
              target=_blank>长安大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.njtu.edu./??&#4;cn" 
                  target=_blank>北方交通大学</A></TD>
                          <TD><A href="http://www.hust.edu.cn/" 
                target=_blank>华中科技大学</A></TD>
                          <TD><A href="http://www.csu.edu.cn/" target=_blank>中南大学</A></TD>
                          <TD><A href="http://www.cufe.edu.cn/" 
                target=_blank>中央财经大学</A></TD>
                        </TR>
                        <TR>
                          <TD height="25"><A href="http://www.whut.edu.cn/" 
                target=_blank>武汉理工大学</A></TD>
                          <TD><A href="http://www.china-visual.com/meiyuan/"
				target=_blank>中央美术学院</A></TD>
                          <TD><A href="http://www.bjucmp.edu.cn/" 
                  target=_blank>北京中医药大学</A></TD>
                          <TD><A href="http://www.ncepu.edu.cn/indexp.htm" 
                  target=_blank>华北电力大学</A></TD>
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
