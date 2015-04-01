<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("user_account")="" then
Response.write"对不起，您还没有登陆，无此权限！"
Response.end
end if
%>
<%
set rs=server.createobject("adodb.recordset")
sql="select sheet.course,sheet.score,sheet.property,sheet.sheet_info,subject.credit,subject.term from user_info inner join (sheet inner join subject on sheet.course = subject.course) on user_info.user_ID = sheet.user_ID where user_account ='"&session("user_account")&"' order by subject.term,sheet.score desc"
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

<html>
<head>
<script language="javascript">
<!--

window.status="欢迎访问南京大学物理系研究生管理信息系统！"
//-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>							南大物理系研究生成绩单</title>
<link href="../style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style12 {font-size: 12px;
	color: #004080;
}
.style17 {color: #FF6633}
.style18 {	font-size: 15px;
	color: #004080;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<div align="center">
  <table width="722"  cellspacing="0" cellpadding="0">
    <%if Not(rs.bof and rs.eof) then
			NumRecord=rs.recordcount
			rs.pagesize=25
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
      <td height="35" bgcolor="#DFE6FF"><div align="center" class="style3">您 的 成 绩 如 下：</div></td>
    </tr>
    <tr>
      <td height="25"><div align="center">
          <table width="720"  border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" class="thin">
            <tr>
              <td width="225" height="24"><div align="center"><span class="style12">课 程</span></div></td>
              <td width="61" height="24"><div align="center" class="style12">
                  <div align="center">成 绩</div>
              </div></td>
              <td width="61" height="24"><div align="center" class="style12">
                  <div align="center">学 分</div>
              </div></td>
              <td width="91"><div align="center" class="style12">
                  <div align="center">课程性质</div>
              </div></td>
              <td width="151"><div align="center" class="style12">
                  <div align="center">修读时间</div>
              </div></td>
              <td width="122"><div align="center" class="style12">
                  <div align="center">备注信息</div>
              </div></td>
            </tr>
            <%if Not(rs.bof and rs.eof) then
	rs.move (Cint(NoncePage)-1)*25,1
	for i=1 to rs.pagesize
%>
            <tr>
              <td height="24"><div align="center" class="style12">
                  <div align="center" class="style12">
                    <div align="center"><%=rs("course")%></div>
                  </div>
              </div></td>
              <td width="61" height="24"><div align="center" class="style12">
                  <div align="center"><%=rs("score")%></div>
              </div></td>
              <td width="61" height="24"><div align="center" class="style12">
                  <div align="center"><%=rs("credit")%></div>
              </div></td>
              <td width="91"><div align="center" class="style12">
                  <div align="center"><%=rs("property")%></div>
              </div></td>
              <td width="151"><div align="center" class="style12">
                  <div align="center"><%=rs("term")%></div>
              </div></td>
              <td width="122"><div align="center" class="style12">
                  <div align="center"><%=HTMLEncode(rs("sheet_info"))%></div>
              </div></td>
            </tr>
            <%rs.movenext
     if rs.eof then exit for
	next
else
	response.write "<tr><td colspan=13><marquee scrolldelay=120 behavior=alternate><font class='style5' color='#ff6633'>暂时没有成绩录入!!!</font></marquee></td></tr>"
end if	
rs.close
set rs=nothing
%>
            <tr>
              <td height="24" colspan="14"><div align="right"> <span class="style12">
                  <input type="hidden" name="page" value="<%=NoncePage%>">
                  <%
if NoncePage>1 then
	response.write "|<a href=user_sheet.asp?page=1>首 页</a>| |<a href=user_sheet.asp?page="&NoncePage-1&">上一页</a>|&nbsp"
else
	response.write "|首 页| |上一页|&nbsp"
end if
if Cint(Trim(NoncePage))<Cint(Trim(NumPage)) then
	response.write "|<a href=user_sheet.asp?page="&NoncePage+1&">下一页</a>| |<a href=user_sheet.asp?page="&NumPage&">尾 页</a>|"
else
	response.write "|下一页| |尾 页|"
end if
%>
&nbsp;页次：<span class="style17"><%=NoncePage%></span>/<span class="style17"><%=NumPage%></span> 共<span class="style17"><%=NumRecord%></span>条记录</span>&nbsp; </div></td>
            </tr>
          </table>
      </div></td>
    </tr>
    <tr bgcolor="#66CCFF">
      <td height="30" colspan="14" bgcolor="#DFE6FF"><div align="center"><a href="sheet_print.asp"><img src="userimages/print.gif" width="70" height="23" border="0" align="absmiddle" onClick="javascript:window.print()"></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="user_sheet.asp"><img src="userimages/return.gif" width="49" height="23" border="0" align="absmiddle"></a> </div></td>
    </tr>
  </table>

</div>
</body>
</html>
