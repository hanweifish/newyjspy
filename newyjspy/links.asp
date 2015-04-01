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
                              <td width="15%" rowspan="2" bgcolor="#F3F7FC"  scope="col"><div align="center" class="style3">校内链接</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://www.nju.edu.cn/">南大首页</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://grawww.nju.edu.cn/">研究生院</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://lib.nju.edu.cn/">图书馆</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://info.nju.edu.cn/">数字化校园</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><SPAN class=main><FONT color=#336699><a href="http://www.nju.edu.cn/cps/site/NJU/njuc/lx.htm">院系联络方法</a></FONT></SPAN></td>
                            </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><a href="http://tuanwei.nju.edu.cn/njuyh/">研究生会</a></td>
                              <td bgcolor="#E6ECF7"><a href="http://www.nju.edu.cn/cps/site/NJU/njuc/timetable/index.htm">列车时刻</a></td>
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
                              <td width="15%" rowspan="2" bgcolor="#E6ECF7"  scope="col"><div align="center" class="style3">就业相关</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://job.nju.edu.cn/">南大就业指导中心</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://www.njbys.com/">南京毕业生就业网</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://www.jsbys.com.cn/index.aspx">江苏毕业生就业网</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://www.js.lm.gov.cn/">中国劳动力市场网</a></td>
                              </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><a href="http://www.firstjob.com.cn/">上海毕业生就业网</a></td>
                              <td bgcolor="#E6ECF7"><a href="http://www.chinahr.com/">中华英才网</a></td>
                              <td bgcolor="#E6ECF7"><a href="http://www.js.lm.gov.cn/">江苏省职业介绍中心</a></td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                              </tr>
                          </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td height="55"><div align="center">
                          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="15%" rowspan="2" bgcolor="#F3F7FC"  scope="col"><div align="center" class="style3">常用搜索</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://www.google.com/">www.google.com</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://www.baidu.com/">百度搜索</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://thephy.nju.edu.cn/">ThePhy搜索</a></td>
                              <td bgcolor="#F3F7FC" scope="col"><a href="http://e.pku.edu.cn/">北大天网搜索</a></td>
                            </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><a href="http://www.sheenk.com/">星空搜索</a></td>
                              <td bgcolor="#E6ECF7"><a href="http://sesa.nju.edu.cn/search/">电子系石头城</a></td>
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
                              <td width="15%" rowspan="2"  scope="col"><div align="center" class="style3">新闻中心</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://news.sina.com.cn/">新浪新闻</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://news.sohu.com/">搜狐新闻</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A href="http://www.xinhuanet.com/">新华社</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://www.peopledaily.com.cn/">人民日报</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://www.chinanews.com.cn/">中新网</A></td>
                            </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><A 
                        href="http://news.tom.com/">TOM新闻</A></td>
                              <td bgcolor="#E6ECF7"><a href="http://www.nanfangdaily.com.cn/zm/20050106/">南方周末</a></td>
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
                              <td width="15%" rowspan="2" bgcolor="#F3F7FC"  scope="col"><div align="center" class="style3">软件下载</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://www.skycn.com/">天空软件</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://www.onlinedown.net/">华军软件园</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://download.21cn.com/">21cn下载</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://download.pchome.net/">电脑之家</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A href="http://www.pconline.com.cn/download/" 
                        target=_blank>太平洋下载</A></td>
                            </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><A 
                        href="http://download.sina.com.cn/">新浪下载</A></td>
                              <td bgcolor="#E6ECF7"><a href="http://dl.163.com/">网易软件下载</a></td>
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
                              <td width="15%" rowspan="2"  scope="col"><div align="center" class="style3">音乐Mp3</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
      href="http://www.mtvtop.com/">中国音乐在线</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
      href="http://game.baidu.com/traf/click.php?id=579&amp;url=http://www.9sky.com/">九天音乐</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
      onmouseover="window.status='';return true" 
      href="http://www.hao123.com/fra/music/tyfocom.htm">天虎音乐网</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
      href="http://www.hao123.com/music/chinayy.htm" target=_blank>中国影音世界</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
      href="http://www.99music.net/new.htm">久久音乐</A></td>
                            </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><A 
      href="http://www.hao123.com/fra/music/chinamp3.htm">音乐极限</A></td>
                              <td bgcolor="#E6ECF7"><A 
      href="http://www.hao123.com/fra/music/soguacom.htm" 
      target=_blank>MP3搜刮网</A></td>
                              <td bgcolor="#E6ECF7"><A 
      href="http://music.feifa.com/">星星音乐谷</A></td>
                              <td bgcolor="#E6ECF7"><A 
      href="http://www.etang.com/music/">亿唐音乐</A></td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                            </tr>
                          </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td height="55"><div align="center">
                          <table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="15%" rowspan="2" bgcolor="#F3F7FC"  scope="col"><div align="center" class="style3">体育频道</div></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://sports.sina.com.cn/">新浪竞技风暴</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://sports.163.com/">网易体育频道</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://sports.sohu.com/" 
                        target=_blank>搜狐体育</A></td>
                              <td bgcolor="#F3F7FC" scope="col"><A 
                        href="http://www.cctv.com/sports/">cctv体育频道</A></td>
                              </tr>
                            <tr>
                              <td bgcolor="#E6ECF7"><A 
                        href="http://sports.tom.com/">TOM鲨威体坛</A></td>
                              <td bgcolor="#E6ECF7"><A 
      href="http://www.beijing-olympic.org.cn/" 
      target=_blank>2008奥运官方</A></td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                              <td bgcolor="#E6ECF7">&nbsp;</td>
                              </tr>
                          </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td bgcolor="#F3F7FC"><div align="center"><a href="university.asp" class="style2">| 教育部直属高校一览 |</a>
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
