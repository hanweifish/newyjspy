<!--#include file = "ConfigUp.asp"-->
<script language="javascript">
window.status="欢迎访问南京大学物理系研究生管理信息系统！"
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="物理系研究生信息管理，南京大学">
<meta name="description" content="南京大学物理系研究生注册系统，参与选课，重要信息发布">
<title>物理系研究生信息管理系统</title>
<link href="style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
-->
</style>
<!--#include file="top1.asp"-->
<table width="700"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="25" valign="bottom" class="style3">重要文件下载</td>
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
                <td width="70"><div align="left">文件名称：</div></td>
                <td><a href="./Upload/<%=fpath%>" target="_NEWwIN"><%=ftitle%>( 文件大小：<%=fsize%> bytes)</a> &nbsp;&nbsp;文件类型：<%=ftype%></td>
                <td align="right"></td>
              </tr>
              <tr class="text">
                <td><div align="right"></div></td>
                <td colspan="2">说明：<%=fdesc%></td>
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
                <td bgcolor="#FFFFFF">尚无新发布内容!</td>
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
		  If Page > 1 Then Response.Write ("<a href='?page=1'>首页</a><a href='?page="& Page - 1 &"'>上一页</a>")
		  If page < pages Then Response.Write ("&nbsp;<a href='?page="& Page + 1 &"'>下一页</a>&nbsp;<a href='?page="& Pages &"'>末页</a>")
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
    <td height="25" valign="bottom" class="style2"> 表 格 下 载</td>
  </tr>
  <tr>
    <td height="30"><table width="100%"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="40" class="style3"><hr></td>
      </tr>
      <tr>
        <td height="40" class="style3">研 究 生 培 养</td>
      </tr>
      <tr>
        <td><a href="http://grawww.nju.edu.cn/yjspy/bspyzh.doc">博士研究生培养计划</a><br><br>
          <a href="http://grawww.nju.edu.cn/yjspy/sspyzh.doc">硕士研究生培养计划</a><br><br>
          <a href="download/jxshx.doc" ><font color="#FF0000">教学实习考核表</font></a><br>
          <br>
          <a href="http://grawww.nju.edu.cn/yjspy/sshyb.doc">师生互选志愿表</a><br>
          <br>
          <a href="http://grawww.nju.edu.cn/yjspy/zbszqkh.doc">直博生中期考核表</a><br><br>          
          <a href="http://grawww.nju.edu.cn/yjsxw/SSLCCL/SSLC.htm">硕士答辩流程及相关材料下载</a><br>
          <br>
          <a href="http://grawww.nju.edu.cn/yjsxw/BSLCCL/BSLC.htm">博士答辩流程及相关材料下载</a></td>
      </tr>
      <tr>
        <td height="40" class="style3"><hr></td>
      </tr>
      <tr>
        <td height="40" class="style3">研 究 生 学 籍 相 关</td>
      </tr>
      <tr>
        <td><br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/hjb.doc">申办户籍证明表</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/qksmbg.doc">情况说明报告</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/lxsxd.doc">离校手续单</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/sqzzy.doc">转专业申请表</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/Blxj.doc">保留学籍申请表</a><br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/Zfcglx.doc"><br>
    自费出国留学申请表</a><br>
    <br>
    <a href="http://grawww.nju.edu.cn/Yjsxj/Tqby.doc">提前毕业申请表</a><br>
    <br>
    <a href="http://grawww.nju.edu.cn/Yjsxj/Yqdb.doc">延期答辩申请表</a></td>
      </tr>
      <tr>
        <td height="40" class="style3"><hr></td>
      </tr>
      <tr>
        <td height="40" class="style3">研 究 生 出 国 管 理</td>
      </tr>
      <tr>
        <td>          <br>
          <a href="http://grawww.nju.edu.cn/yjspy/doctor.doc">博士证书翻译件</a><br>
          <br>
          <a href="http://grawww.nju.edu.cn/yjspy/Master.doc">硕士证书翻译件</a><br>
          <br>
          <a href="http://grawww.nju.edu.cn/yjspy/cjd.doc">外文成绩单</a><br>
          <br>
          <a href="http://grawww.nju.edu.cn/yjspy/hongkong.doc">申请短期赴港工作事宜表</a><br>
		  <br>
          <a href="http://grawww.nju.edu.cn/yjspy/fgsqb.doc">申请赴港事宜表</a><br>
		  <a href="http://grawww.nju.edu.cn/yjspy/cgtjb.doc"><br>
          研究生出国（境）推荐表</a><br>
		  <br>
          <a href="http://wb.nju.edu.cn/download/biao/postgraduate.doc">研究生出国（境）申请表</a><br>
		  <br>
          <a href="http://grawww.nju.edu.cn/yjspy/cgtqb.doc">假期出国（境）探亲、旅游申请表</a><br>
		  <br>
          <a href="http://grawww.nju.edu.cn/yjspy/ZDZM.doc">研究生在读证明</a><br>
		  <br>
          <a href="http://wb.nju.edu.cn/download/biao/postgraduate.doc">研究生出国（境）申请表</a><br>
		  <br>
		  <a href="http://grawww.nju.edu.cn/yjspy/fgsqb.doc">申请赴港事宜表</a></td>
      </tr>
      <tr>
        <td height="40" class="style3"><hr></td>
      </tr>
      <tr>
        <td height="40" class="style3">勤工、贷款、三助、社会实践相关</td>
      </tr>
      <tr>
        <td><br>
            <a href="http://grawww.nju.edu.cn/qgjxdk/sq.doc">勤助岗申请表</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/qgjxdk/ff.doc">津贴发放表</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/yjssc/zyfb.xls">助研费发放表</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/yjssc/szgwsqb.doc">思政助理岗位申请表</a><br>
            <br>
            <a href="http://grawww.nju.edu.cn/Yjsxj/Blxj.doc">保留学籍申请表</a><br>
            <a href="http://grawww.nju.edu.cn/yjssc/ZJSQB.DOC"><br>
    赴疆支教申请表</a></td>
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
