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
                    <td scope="col">&nbsp;</td>
                  </tr>
                  <tr>
                    <td scope="col"><div align="center"><a href="http://grawww.nju.edu.cn/yjsds/gxds.asp?depno=022" class="style3">导师风采</a></div></td>
                  </tr>
                  <tr>
                    <td><div align="center"><a href="yjspy.asp" class="style3">研究生培养</a></div></td>
                  </tr>
                  <tr>
                    <td><div align="center"><a href="yjsxw.asp" class="style3">研究生学位</a></div></td>
                  </tr>
                  <tr>
                    <td><div align="center"><a href="yjsxj.asp" class="style3">研究生学籍</a></div></td>
                  </tr>
                  <tr>
                    <td><div align="center"><a href="yjsjx.asp" class="style2">研究生教学</a></div></td>
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
            <td height="35"><div align="center"><a href="http://www.nju.edu.cn/">南 京 大 学</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="http://physics.nju.edu.cn/">南 大 物 理 系</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="http://bbs.nju.edu.cn/">南 大 小 百 合</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="http://grawww.nju.edu.cn/">研 究 生 院</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="http://job.nju.edu.cn/">就 业 指 导 中 心</a></div></td>
          </tr>
          <tr>
            <td height="35"><div align="center"><a href="links.asp" class="style3">&gt;&gt;&gt; More</a></div></td>
          </tr>
          <tr>
            <td height="35">&nbsp;</td>
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
                  <td height="54" background="user/userimages/titlebk1.gif"><div align="center"><img src="indeximages/rcpy_title.gif" width="523" height="45"></div></td>
                </tr>
                <tr>
                  <td height="53" background="user/userimages/titlebk2.gif">&nbsp;</td>
                </tr>
                <tr>
                  <td height="560" valign="top" background="user/userimages/titlebk.gif"><div align="center"><iframe src="http://grawww.nju.edu.cn/yjsjx/yjsjx.htm" name="denote" width="550" marginwidth="0" height="550" marginheight="0" align="middle" scrolling="auto"  scrolltop=1  frameborder="0" hspace="0" vspace="0" allowtransparency="true"></iframe>
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
