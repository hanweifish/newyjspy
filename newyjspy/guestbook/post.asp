<%@CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="session.asp"-->


<script language="javascript">
	function CheckForm()
	{
		var msg = "";
		if(document.form.forum_title.value == "")
			{
				msg = msg + "      请输入标题!\n\n";
			}
		if(document.form.forum_content.value == "")
			{
				msg = msg + "      请输入内容!\n\n";
			}
		if(msg!="") 
			{
				alert(msg);
				return false;
			}
		return true;
	}
</script>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>留言版</title>
<link href="../style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style8 {font-size: 13px}
-->
</style>
</head>

<body>
<div align="center">
  <!--#include file = "top1.asp"-->
  <table width="840" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="20" background="../images/leftbk.jpg">&nbsp;</td>
      <td colspan="2"><div align="center">
        <table width="85%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="25"><div align="center"></div></td>
          </tr>
          <tr>
            <td height="50"><div align="left"><img src="../images/rotate.gif" width="11" height="11"><span class="style8">&nbsp;留 言 板 &gt;&gt; </span><span class="style1"><a href="post.asp" target="_parent">我要留言</a></span></div></td>
          </tr>
          <tr>
            <td valign="top"><div align="center">
              <form action="post1.asp" method="post" name="form" target="_parent" id="form" onSubmit="return CheckForm();">
                <table width="90%"  border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="183" height="30"><div align="right">留 言 标 题</div></td>
                    <td width="32"><div align="right"></div></td>
                    <td><div align="left">
                      <input name="forum_title" type="text" class="style2" id="forum_title" size="50" maxlength="50">
                    </div></td>
                  </tr>
                  <tr>
                    <td height="200"><div align="right">留 言 内 容</div></td>
                    <td>&nbsp;</td>
                    <td><div align="left">
                      <textarea name="forum_content" cols="60" rows="9" class="style2" id="forum_content"></textarea>
                    </div></td>
                  </tr>
                  <tr>
                    <td height="25" colspan="3"><div align="center">
                      <input name="Submit" type="submit" class="style2" value="发 表">
                      &nbsp;&nbsp;&nbsp;&nbsp;
                      <input name="Submit" type="reset" class="style2"  value="重 填">
                    </div></td>
                    </tr>
                </table>
              </form>
            </div></td>
          </tr>
        </table>
          </div></td>
      <td width="20" background="../images/rightbk.jpg">&nbsp;</td>
    </tr>
  </table>
  
  <!--#include file = "bottom1.asp"-->
</div>
</body>
</html>
