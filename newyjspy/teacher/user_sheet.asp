<!--#include file="conn.asp"-->
<%
if request.cookies("status")="" then
    Response.write"对不起，您还没有登陆，无此权限！"
	Response.end
end if
%>

<%
if session("admin_account")="" or session("user_group")="subadmin" then
Response.write"对不起，您还没有登陆或者无此权限！"
Response.end
end if
%>

<%
	dim fyft, fyst, syft, syst, tyft, tyst, user_account,usernumber,usernumber1,usernumber2
	fyft = "0"
	fyst = "0"
	syft = "0"
	syst = "0"
	tyft = "0"
	tyst = "0"
	user_account = trim(request("user_account"))
	set rsUser = Server.CreateObject("Adodb.Recordset")
	sql_rsUser = "select * from user_info where user_account = '"&user_account&"'"
	rsUser.Open sql_rsUser,conn,1,1
	usernumber=rsUser("user_number")
	usernumber1=left(usernumber,4)
	usernumber2=right(usernumber1,2)
	
	if rsUser.RecordCount = 0 then
	Response.Write ("<script> alert('请重新登陆!');parent.window.history.go(-1)</script>")
	end if
	
	set rsSheetFirstCom = Server.CreateObject("Adodb.Recordset")
	sql_rsSheetFirstCom = "select sheet.course, sheet.score, sheet.sheet_credit, subject.tutor, subject.term, subject.teachway from sheet inner join subject on sheet.course_ID = subject.course_ID where property = '必修课' and year = '第一学年' and user_ID = "&rsUser("user_ID")&" order by subject.term"
	rsSheetFirstCom.Open sql_rsSheetFirstCom,conn,1,1
	
	set rsSheetSecondCom = Server.CreateObject("Adodb.Recordset")
	sql_rsSheetSecondCom = "select sheet.course, sheet.score, sheet.sheet_credit, subject.tutor,  subject.term, subject.teachway from sheet inner join subject on sheet.course_ID = subject.course_ID where property = '必修课' and year = '第二学年' and user_ID = "&rsUser("user_ID")&" order by subject.term"
	rsSheetSecondCom.Open sql_rsSheetSecondCom,conn,1,1
	
	set rsSheetThirdCom = Server.CreateObject("Adodb.Recordset")
	sql_rsSheetThirdCom = "select sheet.course, sheet.score, sheet.sheet_credit, subject.tutor,  subject.term, subject.teachway from sheet inner join subject on sheet.course_ID = subject.course_ID where property = '必修课' and year = '第三学年' and user_ID = "&rsUser("user_ID")&" order by subject.term"
	rsSheetThirdCom.Open sql_rsSheetThirdCom,conn,1,1
	
	set rsSheetFirstOpt = Server.CreateObject("Adodb.Recordset")
	sql_rsSheetFirstOpt = "select sheet.course, sheet.score, sheet.sheet_credit, subject.tutor,  subject.term, subject.teachway from sheet inner join subject on sheet.course_ID = subject.course_ID where property = '选修课' and year = '第一学年' and user_ID = "&rsUser("user_ID")&" order by subject.term"
	rsSheetFirstOpt.Open sql_rsSheetFirstOpt,conn,1,1
	
	set rsSheetSecondOpt = Server.CreateObject("Adodb.Recordset")
	sql_rsSheetSecondOpt = "select sheet.course, sheet.score, sheet.sheet_credit, subject.tutor,  subject.term, subject.teachway from sheet inner join subject on sheet.course_ID = subject.course_ID where property = '选修课' and year = '第二学年' and user_ID = "&rsUser("user_ID")&" order by subject.term"
	rsSheetSecondOpt.Open sql_rsSheetSecondOpt,conn,1,1
	
	set rsSheetThirdOpt = Server.CreateObject("Adodb.Recordset")
	sql_rsSheetThirdOpt = "select sheet.course, sheet.score, sheet.sheet_credit, subject.tutor,  subject.term, subject.teachway from sheet inner join subject on sheet.course_ID = subject.course_ID where property = '选修课' and year = '第三学年' and user_ID = "&rsUser("user_ID")&" order by subject.term"
	rsSheetThirdOpt.Open sql_rsSheetThirdOpt,conn,1,1
	
	set rsBiyeInfo = Server.CreateObject("Adodb.Recordset")
	sql_rsBiyeInfo = "select * from biyeInfo where user_number = '"&rsUser("user_number")&"'"
	rsBiyeInfo.Open sql_rsBiyeInfo,conn,3,3
	
	if (rsBiyeInfo.RecordCount = 0) then
		rsBiyeInfo.AddNew
		rsBiyeInfo("user_Number")=rsUser("user_number")
		rsBiyeInfo("jxsx")=""
		rsBiyeInfo("bylw")=""
		rsBiyeInfo("zdjs")=""
		rsBiyeInfo("dbsj")=""
		rsBiyeInfo("dbjg")=""
		rsBiyeInfo("byrq")=""
		rsBiyeInfo("wphm")=""
		rsBiyeInfo("xwhm")=""
		rsBiyeInfo("fzrqm")=""
		rsBiyeInfo("jwyqm")=""
		rsBiyeInfo.Update		
	end if
	
	set rsJiangcheng = Server.CreateObject("Adodb.Recordset")
	sql_rsJiangcheng = "select * from jiangcheng where user_number = '"&rsUser("user_number")&"'"
	rsJiangcheng.Open sql_rsJiangcheng,conn,1,1
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script language="javascript">
<!--
window.status="欢迎访问南京大学物理系研究生管理信息系统！"
//-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>研究生成绩表</title>
<style type="text/css">
<!--
body,td,th {
	font-size: 13px;
}
body {
	margin-top: 0px;
}
.STYLE1 {
	font-size: 28px;
	font-weight: bold;
}
.STYLE2 {font-size: 16px}
-->
</style></head>

<body>
<table width = "1560" border = "0" cellpadding = "0" cellspacing = "0">
 <tr>
   <td height="20">&nbsp;&nbsp;&nbsp; <a href="admin_index.asp">返 回</a></td>
 </tr>
</table>
<table width="1560" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td colspan="2" rowspan="2"><div align="left"> &nbsp;学号：<%=rsUser("user_number")%></div></td>
    <td colspan="6" rowspan="2"><div align="left">&nbsp;姓名：<%=rsUser("user_name")%></div></td>
    <td colspan="11" rowspan="2"><div align="center" class="STYLE1">南京大学硕士生研究生学习成绩表</div></td>
    <td colspan="3" rowspan="2"><div align="center" class="STYLE2">物 理 系</div></td>
    <td height="40" colspan="5">&nbsp;专业：<%=rsUser("user_major")%></td>
  </tr>
  
  <tr>
    <td height="40" colspan="5">&nbsp;指导教师：<%=rsUser("user_tutor")%></td>
  </tr>
 <%
  if usernumber2=06 then response.Write ("<tr><td width='80' rowspan='3' background='user/userimages/xncjlb.jpg'>&nbsp;</td><td height='40' colspan='8'><div align='center'>第  一  学  年<br/>（自2006年9月至2007年6月）</div></td><td width='80' rowspan='3' background='user/userimages/xncjlb.jpg'>&nbsp;</td><td colspan='8'><div align='center'>第  二  学  年<br/>               （自2007年9月至2008年6月）</div></td><td width='80' rowspan='3' background='user/userimages/xncjlb.jpg'>&nbsp;</td><td colspan='8'><div align='center'>第  三  学  年<br/>（自2008年9月至2009年6月）</div></td></tr> ") end if 
  
   if usernumber2=05 then response.Write "<tr><td width='80' rowspan='3' background='user/userimages/xncjlb.jpg'>&nbsp;</td><td height='40' colspan='8'><div align='center'>第  一  学  年<br/>（自2005年9月至2006年6月）</div></td><td width='80' rowspan='3' background='user/userimages/xncjlb.jpg'>&nbsp;</td><td colspan='8'><div align='center'>第  二  学  年<br/>               （自2006年9月至2007年6月）</div></td><td width='80' rowspan='3' background='user/userimages/xncjlb.jpg'>&nbsp;</td><td colspan='8'><div align='center'>第  三  学  年<br/>（自2007年9月至2008年6月）</div></td></tr> " end if 
  
   if usernumber2=04 then response.Write "<tr><td width='80' rowspan='3' background='user/userimages/xncjlb.jpg'>&nbsp;</td><td height='40' colspan='8'><div align='center'>第  一  学  年<br/>（自2004年9月至2005年6月）</div></td><td width='80' rowspan='3' background='user/userimages/xncjlb.jpg'>&nbsp;</td><td colspan='8'><div align='center'>第  二  学  年<br/>               （自2005年9月至2006年6月）</div></td><td width='80' rowspan='3' background='user/userimages/xncjlb.jpg'>&nbsp;</td><td colspan='8'><div align='center'>第  三  学  年<br/>（自2006年9月至2007年6月）</div></td></tr> " end if 
  
 %> 
  <tr>
    <td width="120" rowspan="2" nowrap="nowrap"><div align="center">课程名称</div></td>
    <td height="20" colspan="3" nowrap="nowrap"><div align="center">上学期</div></td>
    <td colspan="3" nowrap="nowrap"><div align="center">下学期</div></td>
    <td width="80" rowspan="2" nowrap="nowrap"><div align="center">任课教师</div></td>
    <td width="120" rowspan="2" nowrap="nowrap"><div align="center">课程名称</div></td>
    <td height="20" colspan="3" nowrap="nowrap"><div align="center">上学期</div></td>
    <td colspan="3" nowrap="nowrap"><div align="center">下学期</div></td>
    <td width="80" rowspan="2" nowrap="nowrap"><div align="center">任课教师</div></td>
    <td width="120" rowspan="2" nowrap="nowrap"><div align="center">课程名称</div></td>
    <td height="20" colspan="3" nowrap="nowrap"><div align="center">上学期</div></td>
    <td colspan="3" nowrap="nowrap"><div align="center">下学期</div></td>
    <td width="80" rowspan="2" nowrap="nowrap"><div align="center">任课教师</div></td>
  </tr>
  <tr>
    <td width="40" height="40"><div align="center">学分</div></td>
    <td width="40"><div align="center">成绩</div></td>
    <td width="40"><div align="center">教学<br />
    形式</div></td>
    <td width="40"><div align="center">学分</div></td>
    <td width="40"><div align="center">成绩</div></td>
    <td width="40"><div align="center">教学形式</div></td>
    <td width="40" height="40"><div align="center">学分</div></td>
    <td width="40"><div align="center">成绩</div></td>
    <td width="40"><div align="center">教学<br />
      形式</div></td>
    <td width="40"><div align="center">学分</div></td>
    <td width="40"><div align="center">成绩</div></td>
    <td width="40"><div align="center">教学形式</div></td>
    <td width="40" height="40"><div align="center">学分</div></td>
    <td width="40"><div align="center">成绩</div></td>
    <td width="40"><div align="center">教学<br />
      形式</div></td>
    <td width="40"><div align="center">学分</div></td>
    <td width="40"><div align="center">成绩</div></td>
    <td width="40"><div align="center">教学形式</div></td>
  </tr>
  
  <tr>
    <td rowspan="10"><div align="center">
      <p>必</p>
      <p>修</p>
      <p>课</p>
    </div></td>
	<%
		if not (rsSheetFirstCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstCom("course")%></div></td>	<%
		if rsSheetFirstCom("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
	<%
		end if
	%>
    <td><div align="center">&nbsp;<%=rsSheetFirstCom("tutor")%></div></td>
	<%
		rsSheetFirstCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
	
    <td rowspan="10"><div align="center">
      <p>必</p>
      <p>修</p>
      <p>课</p>
    </div></td>
	
	<%
		if not (rsSheetSecondCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondCom("course")%></div></td>	<%
		if rsSheetSecondCom("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
	<%
		end if
	%>
    <td><div align="center">&nbsp;<%=rsSheetSecondCom("tutor")%></div></td>	<%
		rsSheetSecondCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<td rowspan="10"><div align="center">
      <p>必</p>
      <p>修</p>
      <p>课</p>
    </div></td>
   	<%
		if not (rsSheetThirdCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdCom("course")%></div></td>	<%
		if rsSheetThirdCom("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
	<%
		end if
	%>
    <td><div align="center">&nbsp;<%=rsSheetThirdCom("tutor")%></div></td>	<%
		rsSheetThirdCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>

	<%
		if not (rsSheetFirstCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstCom("course")%></div></td>	<%
		if rsSheetFirstCom("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstCom("tutor")%></div></td>	<%
		rsSheetFirstCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondCom("course")%></div></td>	<%
		if rsSheetSecondCom("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondCom("tutor")%></div></td>	<%
		rsSheetSecondCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdCom("course")%></div></td>	<%
		if rsSheetThirdCom("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdCom("tutor")%></div></td>	<%
		rsSheetThirdCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>

	<%
		if not (rsSheetFirstCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstCom("course")%></div></td>	<%
		if rsSheetFirstCom("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstCom("tutor")%></div></td>	<%
		rsSheetFirstCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondCom("course")%></div></td>	<%
		if rsSheetSecondCom("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondCom("tutor")%></div></td>	<%
		rsSheetSecondCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdCom("course")%></div></td>	<%
		if rsSheetThirdCom("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdCom("tutor")%></div></td>	<%
		rsSheetThirdCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
	<%
		if not (rsSheetFirstCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstCom("course")%></div></td>	<%
		if rsSheetFirstCom("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstCom("tutor")%></div></td>	<%
		rsSheetFirstCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondCom("course")%></div></td>	<%
		if rsSheetSecondCom("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondCom("tutor")%></div></td>	<%
		rsSheetSecondCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdCom("course")%></div></td>	<%
		if rsSheetThirdCom("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdCom("tutor")%></div></td>	<%
		rsSheetThirdCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
	<%
		if not (rsSheetFirstCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstCom("course")%></div></td>	<%
		if rsSheetFirstCom("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstCom("tutor")%></div></td>	<%
		rsSheetFirstCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondCom("course")%></div></td>	<%
		if rsSheetSecondCom("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondCom("tutor")%></div></td>	<%
		rsSheetSecondCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdCom("course")%></div></td>	<%
		if rsSheetThirdCom("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdCom("tutor")%></div></td>	<%
		rsSheetThirdCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
	<%
		if not (rsSheetFirstCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstCom("course")%></div></td>	<%
		if rsSheetFirstCom("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstCom("tutor")%></div></td>	<%
		rsSheetFirstCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondCom("course")%></div></td>	<%
		if rsSheetSecondCom("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondCom("tutor")%></div></td>	<%
		rsSheetSecondCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdCom("course")%></div></td>	<%
		if rsSheetThirdCom("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdCom("tutor")%></div></td>	<%
		rsSheetThirdCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
	<%
		if not (rsSheetFirstCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstCom("course")%></div></td>	<%
		if rsSheetFirstCom("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstCom("tutor")%></div></td>	<%
		rsSheetFirstCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondCom("course")%></div></td>	<%
		if rsSheetSecondCom("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondCom("tutor")%></div></td>	<%
		rsSheetSecondCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdCom("course")%></div></td>	<%
		if rsSheetThirdCom("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdCom("tutor")%></div></td>	<%
		rsSheetThirdCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
	<%
		if not (rsSheetFirstCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstCom("course")%></div></td>	<%
		if rsSheetFirstCom("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstCom("tutor")%></div></td>	<%
		rsSheetFirstCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondCom("course")%></div></td>	<%
		if rsSheetSecondCom("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondCom("tutor")%></div></td>	<%
		rsSheetSecondCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdCom("course")%></div></td>	<%
		if rsSheetThirdCom("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdCom("tutor")%></div></td>	<%
		rsSheetThirdCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
	<%
		if not (rsSheetFirstCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstCom("course")%></div></td>	<%
		if rsSheetFirstCom("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstCom("tutor")%></div></td>	<%
		rsSheetFirstCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondCom("course")%></div></td>	<%
		if rsSheetSecondCom("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondCom("tutor")%></div></td>	<%
		rsSheetSecondCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdCom("course")%></div></td>	<%
		if rsSheetThirdCom("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdCom("tutor")%></div></td>	<%
		rsSheetThirdCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
	<%
		if not (rsSheetFirstCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstCom("course")%></div></td>	<%
		if rsSheetFirstCom("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstCom("tutor")%></div></td>	<%
		rsSheetFirstCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondCom("course")%></div></td>	<%
		if rsSheetSecondCom("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondCom("tutor")%></div></td>	<%
		rsSheetSecondCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdCom.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdCom("course")%></div></td>	<%
		if rsSheetThirdCom("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdCom("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdCom("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdCom("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdCom("tutor")%></div></td>	<%
		rsSheetThirdCom.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
    <td rowspan="6"><div align="center">
      <p>选</p>
      <p>修</p>
      <p>课</p>
    </div></td>
	
	<%
		if not (rsSheetFirstOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstOpt("course")%></div></td>	<%
		if rsSheetFirstOpt("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("teachway")%></div></td>
	<%
		end if
	%>
    <td><div align="center">&nbsp;<%=rsSheetFirstOpt("tutor")%></div></td>	<%
		rsSheetFirstOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<td rowspan="6"><div align="center">
      <p>选</p>
      <p>修</p>
      <p>课</p>
    </div></td>
        
	<%
		if not (rsSheetSecondOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondOpt("course")%></div></td>	<%
		if rsSheetSecondOpt("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("teachway")%></div></td>
	<%
		end if
	%>
    <td><div align="center">&nbsp;<%=rsSheetSecondOpt("tutor")%></div></td>	<%
		rsSheetSecondOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<td rowspan="6"><div align="center">
      <p>选</p>
      <p>修</p>
      <p>课</p>
    </div></td>
   	<%
		if not (rsSheetThirdOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdOpt("course")%></div></td>	<%
		if rsSheetThirdOpt("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("teachway")%></div></td>
	<%
		end if
	%>
    <td><div align="center">&nbsp;<%=rsSheetThirdOpt("tutor")%></div></td>	<%
		rsSheetThirdOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
	<%
		if not (rsSheetFirstOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstOpt("course")%></div></td>	<%
		if rsSheetFirstOpt("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstOpt("tutor")%></div></td>	<%
		rsSheetFirstOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondOpt("course")%></div></td>	<%
		if rsSheetSecondOpt("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondOpt("tutor")%></div></td>	<%
		rsSheetSecondOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdOpt("course")%></div></td>	<%
		if rsSheetThirdOpt("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdOpt("tutor")%></div></td>	<%
		rsSheetThirdOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
	<%
		if not (rsSheetFirstOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstOpt("course")%></div></td>	<%
		if rsSheetFirstOpt("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstOpt("tutor")%></div></td>	<%
		rsSheetFirstOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondOpt("course")%></div></td>	<%
		if rsSheetSecondOpt("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondOpt("tutor")%></div></td>	<%
		rsSheetSecondOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdOpt("course")%></div></td>	<%
		if rsSheetThirdOpt("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdOpt("tutor")%></div></td>	<%
		rsSheetThirdOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
	<%
		if not (rsSheetFirstOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstOpt("course")%></div></td>	<%
		if rsSheetFirstOpt("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstOpt("tutor")%></div></td>	<%
		rsSheetFirstOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondOpt("course")%></div></td>	<%
		if rsSheetSecondOpt("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondOpt("tutor")%></div></td>	<%
		rsSheetSecondOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdOpt("course")%></div></td>	<%
		if rsSheetThirdOpt("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdOpt("tutor")%></div></td>	<%
		rsSheetThirdOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
	<%
		if not (rsSheetFirstOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstOpt("course")%></div></td>	<%
		if rsSheetFirstOpt("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstOpt("tutor")%></div></td>	<%
		rsSheetFirstOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondOpt("course")%></div></td>	<%
		if rsSheetSecondOpt("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondOpt("tutor")%></div></td>	<%
		rsSheetSecondOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdOpt("course")%></div></td>	<%
		if rsSheetThirdOpt("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdOpt("tutor")%></div></td>	<%
		rsSheetThirdOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
	<%
		if not (rsSheetFirstOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetFirstOpt("course")%></div></td>	<%
		if rsSheetFirstOpt("term") = "上学期" then
		fyft = CInt(fyft)+CInt(rsSheetFirstOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetFirstOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		fyst = CInt(fyst)+CInt(rsSheetFirstOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetFirstOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetFirstOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetFirstOpt("tutor")%></div></td>	<%
		rsSheetFirstOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetSecondOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetSecondOpt("course")%></div></td>	<%
		if rsSheetSecondOpt("term") = "上学期" then
		syft = CInt(syft)+CInt(rsSheetSecondOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetSecondOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		syst = CInt(syst)+CInt(rsSheetSecondOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetSecondOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetSecondOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetSecondOpt("tutor")%></div></td>	<%
		rsSheetSecondOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
<%
		if not (rsSheetThirdOpt.eof ) then
	%>
    <td height="20"><div align="left"><%=rsSheetThirdOpt("course")%></div></td>	<%
		if rsSheetThirdOpt("term") = "上学期" then
		tyft = CInt(tyft)+CInt(rsSheetThirdOpt("sheet_credit"))
	%>
    <td><div align="center"><%=rsSheetThirdOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("teachway")%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		else
		tyst = CInt(tyst)+CInt(rsSheetThirdOpt("sheet_credit"))
	%>
<td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=rsSheetThirdOpt("sheet_credit")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("score")%></div></td>
    <td><div align="center"><%=rsSheetThirdOpt("teachway")%></div></td>
	<%
		end if
	%>
<td><div align="center">&nbsp;<%=rsSheetThirdOpt("tutor")%></div></td>	<%
		rsSheetThirdOpt.MoveNext
		else
	%>
    <td height="20"><div align="left">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
</tr>
  <tr>
    <td rowspan="2"><div align="center">累计学分</div></td>
    <td height="20"><div align="center">学期合计</div></td>
    <td height="20" colspan="3"><div align="center"><%=fyft%></div></td>
    <td height="20" colspan="3"><div align="center"><%=fyst%></div></td>
    <td height="20">&nbsp;</td>
    <td height="20">&nbsp;</td>
    <td height="20">&nbsp;</td>
    <td height="20" colspan="3"><div align="center"><%=syft%></div></td>
    <td height="20" colspan="3"><div align="center"><%=syst%></div></td>
    <td height="20">&nbsp;</td>
    <td height="20">&nbsp;</td>
    <td height="20">&nbsp;</td>
    <td height="20" colspan="3"><div align="center"><%=tyft%></div></td>
    <td height="20" colspan="3"><div align="center"><%=tyst%></div></td>
    <td height="20">&nbsp;</td>
  </tr>
  <tr>
    <td height="20"><div align="center">学年合计</div></td>
    <td height="20" colspan="7"><div align="center"><%=CInt(fyft) + CInt(fyst)%></div></td>
    <td height="20" colspan="9"><div align="center"><%=CInt(syft) + CInt(syst)%></div></td>
    <td height="20" colspan="9"><div align="center"><%=CInt(tyft) + CInt(tyst)%></div></td>
  </tr>
  <tr>
    <td><div align="center">教学实习</div></td>
    <td height="40" colspan="26"><div align="left">&nbsp;&nbsp;<%=rsBiyeInfo("jxsx")%></div></td>
  </tr>
  
  <tr>
    <td><div align="center">休、退<br />
      学及奖<br />
      惩情况</div></td>
    <td colspan="26"><div align="center">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
	  	
		<%
			if rsJiangcheng.RecordCount = 0 then
		%>
        <tr>
          <td height="20"><div align="left">&nbsp;&nbsp;无</div></td>
        </tr>
		<%
			end if		
		%>
		<%
	  		if rsJiangcheng.RecordCount <> 0 then
			for i=1 to rsJiangcheng.RecordCount 
	  	%>
        <tr>
          <td height="20"><div align="left">&nbsp;&nbsp;<%=rsJiangcheng("record")%></div></td>
        </tr>
		<%
			rsJiangcheng.MoveNext
			next
			end if
		%>
      </table>
    </div></td>
  </tr>
  
  <tr>
    <td rowspan="2"><div align="center">毕业论<br />
    文题目</div></td>
    <td colspan="8" rowspan="2"><div align="center">&nbsp;<%=rsBiyeInfo("bylw")%></div></td>
    <td colspan="2" rowspan="2"><div align="center">论文指导教师姓名</div></td>
    <td colspan="7" rowspan="2"><div align="center">&nbsp;<%=rsBiyeInfo("zdjs")%></div></td>
    <td height="20" colspan="2"><div align="center">答辩时间</div></td>
    <td height="20" colspan="7"><div align="center">&nbsp;<%=rsBiyeInfo("dbsj")%></div></td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">答辩结果</div></td>
    <td height="20" colspan="7"><div align="center">&nbsp;<%=rsBiyeInfo("dbjg")%></div></td>
  </tr>
  <tr>
    <td rowspan="2"><div align="center">毕业<br />
    记录</div></td>
    <td colspan="8" rowspan="2"><div align="center">毕业日期:<%=rsBiyeInfo("byrq")%></div></td>
    <td height="40" colspan="2"><div align="center">系负责人签名</div></td>
    <td colspan="7"><div align="center">&nbsp;<%=rsBiyeInfo("fzrqm")%></div></td>
    <td colspan="2"><div align="center">毕业文凭号码</div></td>
    <td colspan="7"><div align="center">&nbsp;<%=rsBiyeInfo("wphm")%></div></td>
  </tr>
  
  <tr>
    <td height="40" colspan="2"><div align="center">系教务员签名</div></td>
    <td colspan="7"><div align="center">&nbsp;<%=rsBiyeInfo("jwyqm")%></div></td>
    <td colspan="2"><div align="center">学位证号码</div></td>
    <td colspan="7"><div align="center">&nbsp;<%=rsBiyeInfo("xwhm")%></div></td>
  </tr>
</table>
<%
	rsUser.Close
	rsJiangcheng.Close
	rsBiyeInfo.Close
	rsSheetFirstCom.Close
	rsSheetSecondCom.Close
	rsSheetThirdCom.Close
	rsSheetFirstOpt.Close
	rsSheetSecondOpt.Close
	rsSheetThirdOpt.Close
%>
</body>
</html>
