<!--#include file = "ConfigUp.asp"-->
<script language="javascript">
window.status="��ӭ�����Ͼ���ѧ����ϵ�о���������Ϣϵͳ��"
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="����ϵ�о�����Ϣ�����Ͼ���ѧ">
<meta name="description" content="�Ͼ���ѧ����ϵ�о���ע��ϵͳ������ѡ�Σ���Ҫ��Ϣ����">
<title>����ϵ�о�����Ϣ����ϵͳ</title>
<link href="style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
-->
</style>
<!--#include file="top1.asp"-->
<table width="700"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="25" valign="bottom" class="style3">��Ҫ�ļ�����</td>
  </tr>
  <tr>
    <td height="40" class="style2"><hr></td>
  </tr>
  <tr>
    <td class="style2"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
      <tr>
        <td height="25">
		<%
					set frst = Server.CreateObject("adodb.recordset")
					sql = "select * from Fupload order by Uploadtime desc"
					frst.open sql,myconn,1,1
					fcount = frst.recordcount
					if fcount > 0 then 	
						dim tbbgcolor
						dim maxperpage,pages,page
						maxperpage = 4
						frst.pagesize = maxperpage
						pages = frst.pagecount
						page = Request.QueryString("page")
						if not isnumeric(page) then page = 1 else page = cint(page)
						if page < 1 then page = 1
						if page > pages then page = pages
						frst.absolutepage = page
						for i = 1 to maxperpage
							if frst.eof then exit for
							if i mod 2 = 0 then tbbgcolor = "" else tbbgcolor = "#0066cc"
							fid = frst("Fupload_ID").Value
							ftitle = frst("FileTitle").Value
							fdesc = frst("FileDesc").Value
							ftype = frst("FileType").Value
							fpath = frst("FilePath").Value
							fsize = frst("Filesize").Value
							fuploadtime = frst("uploadTime").Value
					%>
            <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
              <tr class="text">
                <td width="70"><div align="left">�ļ����ƣ�</div></td>
                <td><a href="./Upload/<%=fpath%>" target="_NEWwIN"><%=ftitle%>( �ļ���С��<%=fsize%> bytes)</a> &nbsp;&nbsp;�ļ����ͣ�<%=ftype%></td>
                <td align="right"></td>
              </tr>
              <tr class="text">
                <td><div align="right"></div></td>
                <td colspan="2">˵����<%=fdesc%></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="1"></td>
                <td height="1" colspan="2"></td>
              </tr>
            </table>
            <%
		  	frst.movenext
		next
	  else
	  %>
            <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
              <tr class="text">
                <td bgcolor="#FFFFFF">�����·�������!</td>
              </tr>
            </table>
            <%
	  end if
	  frst.close
	  set frst = nothing
	  myconn.close
	  set myconn = nothing
	  %>
            <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
              <tr class="text">
                <td align="center"><%
		  If Page > 1 Then Response.Write ("<a href='?page=1'>��ҳ</a><a href='?page="& Page - 1 &"'>��һҳ</a>")
		  If page < pages Then Response.Write ("&nbsp;<a href='?page="& Page + 1 &"'>��һҳ</a>&nbsp;<a href='?page="& Pages &"'>ĩҳ</a>")
		  %></td>
              </tr>
          </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="25" class="style2"><hr></td>
  </tr>
  <tr>
    <td height="25" valign="bottom" class="style2"> �� �� �� ��</td>
  </tr>
  <tr>
    <td height="30"><table width="100%"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="40" class="style3"><hr></td>
      </tr>
      <tr>
        <td height="40" class="style3">�� �� �� �� ��</td>
      </tr>
      <tr>
        <td><a href="http://grawww.nju.edu.cn/yjspy/bspyzh.doc">��ʿ�о��������ƻ�</a><br><br>
          <a href="http://grawww.nju.edu.cn/yjspy/sspyzh.doc">˶ʿ�о��������ƻ�</a><br><br>
          <a href="download/jxshx.doc" ><font color="#FF0000">��ѧʵϰ���˱�</font></a><br>
          <br>
          <a href="http://grawww.nju.edu.cn/yjspy/sshyb.doc">ʦ����ѡ־Ը��</a><br>
          <br>
          <a href="http://grawww.nju.edu.cn/yjspy/zbszqkh.doc">ֱ�������ڿ��˱�</a><br><br>          
          <a href="http://grawww.nju.edu.cn/yjsxw/SSLCCL/SSLC.htm">˶ʿ������̼���ز�������</a><br>
          <br>
          <a href="http://grawww.nju.edu.cn/yjsxw/BSLCCL/BSLC.htm">��ʿ������̼���ز�������</a></td>
      </tr>
      <tr>
        <td height="40" class="style3"><hr></td>
      </tr>
      <tr>
        <td height="40" class="style3">�� �� �� ѧ �� �� ��</td>
      </tr>
      <tr>
        <td><br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/hjb.doc">��컧��֤����</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/qksmbg.doc">���˵������</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/lxsxd.doc">��У������</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/sqzzy.doc">תרҵ�����</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/Blxj.doc">����ѧ�������</a><br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/Zfcglx.doc"><br>
    �Էѳ�����ѧ�����</a><br>
    <br>
    <a href="http://grawww.nju.edu.cn/Yjsxj/Tqby.doc">��ǰ��ҵ�����</a><br>
    <br>
    <a href="http://grawww.nju.edu.cn/Yjsxj/Yqdb.doc">���ڴ�������</a></td>
      </tr>
      <tr>
        <td height="40" class="style3"><hr></td>
      </tr>
      <tr>
        <td height="40" class="style3">�� �� �� �� �� �� ��</td>
      </tr>
      <tr>
        <td>          <br>
          <a href="http://grawww.nju.edu.cn/yjspy/doctor.doc">��ʿ֤�鷭���</a><br>
          <br>
          <a href="http://grawww.nju.edu.cn/yjspy/Master.doc">˶ʿ֤�鷭���</a><br>
          <br>
          <a href="http://grawww.nju.edu.cn/yjspy/cjd.doc">���ĳɼ���</a><br>
          <br>
          <a href="http://grawww.nju.edu.cn/yjspy/hongkong.doc">������ڸ��۹������˱�</a><br>
		  <br>
          <a href="http://grawww.nju.edu.cn/yjspy/fgsqb.doc">���븰�����˱�</a><br>
		  <a href="http://grawww.nju.edu.cn/yjspy/cgtjb.doc"><br>
          �о��������������Ƽ���</a><br>
		  <br>
          <a href="http://wb.nju.edu.cn/download/biao/postgraduate.doc">�о������������������</a><br>
		  <br>
          <a href="http://grawww.nju.edu.cn/yjspy/cgtqb.doc">���ڳ���������̽�ס����������</a><br>
		  <br>
          <a href="http://grawww.nju.edu.cn/yjspy/ZDZM.doc">�о����ڶ�֤��</a><br>
		  <br>
          <a href="http://wb.nju.edu.cn/download/biao/postgraduate.doc">�о������������������</a><br>
		  <br>
		  <a href="http://grawww.nju.edu.cn/yjspy/fgsqb.doc">���븰�����˱�</a></td>
      </tr>
      <tr>
        <td height="40" class="style3"><hr></td>
      </tr>
      <tr>
        <td height="40" class="style3">�ڹ���������������ʵ�����</td>
      </tr>
      <tr>
        <td><br>
            <a href="http://grawww.nju.edu.cn/qgjxdk/sq.doc">�����������</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/qgjxdk/ff.doc">�������ű�</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/yjssc/zyfb.xls">���зѷ��ű�</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/yjssc/szgwsqb.doc">˼�������λ�����</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/Blxj.doc">����ѧ�������</a><br>
            <a href="http://grawww.nju.edu.cn/yjssc/ZJSQB.DOC"><br>
    ����֧�������</a></td>
      </tr>
      <tr>
        <td>&nbsp;
          </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
</table>
<!--#include file="bottom.asp"-->
</html>
