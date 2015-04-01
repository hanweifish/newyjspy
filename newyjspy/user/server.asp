<%
	strSourceFile = Server.MapPath("../inc/config.xml")
	Set objXML = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
	objXML.load(strSourceFile)
	Set objRoot = objXML.selectSingleNode("Config")
	dim info,info1
	info=objRoot.childNodes.item(0).text
	info1=objRoot.childNodes.item(1).text
%>
<table width="603" height="236" border="0" cellpadding="0" cellspacing="0" background="../includeimages/service.gif">
  <tr>
    <td width="406" valign="top"><div align="center">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="60" valign="middle"><div align="center"><img src="../includeimages/phone.gif" width="150" height="35"></div></td>
        </tr>
        <tr>
          <td valign="top"><div align="center">
            <table width="90%" height="90%"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="54%" height="100" class="style2"><div align="left">
                    <p>&nbsp;&nbsp;&nbsp;&nbsp;<%=info%></p>
                    <p>&nbsp;&nbsp;&nbsp;&nbsp;<%=info1%></p>
                  </div>
                    <div align="left"></div></td>
              </tr>
            </table>
          </div></td>
        </tr>
      </table>
      </div>      <div valign="top" align="center">
      </div>      <div align="center"></div></td>
    <td width="3"><img src="../includeimages/bar1.gif" width="3" height="146"></td>
    <td width="192" valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="200" height="60" valign="middle"><div align="center"><a href="../Site_login.asp"><img src="../includeimages/stusite.gif" width="150" height="35" border="0"></a></div></td>
        </tr>
        <tr>
          <td width="200" height="176" valign="top"><div align="center">
              <table width="150" height="60%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="5" colspan="2" >&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="middle"><div align="center">
                      <marquee behavior="scroll" align="absmiddle" direction="right" width="120" height="25" loop="-1" scrollamount="80" scrolldelay="1000" onMouseOver="javascript: this.stop()" onMouseOut="javascript:this.start()">
                      <%	set rsSite2=server.createobject("adodb.recordset")
								sitesql="select * from user_site order by site_name"
								rsSite2.open sitesql,conn,1,1
							%>
                      <%if not rsSite2.eof then
							for i=0 to rsSite2.recordcount
							%>
                      <a class="style2" href="<%=rsSite2("site_url")%>" title="<%=rsSite2("site_admin")%> : <%=rsSite2("site_info")%>"><%=rsSite2("site_name")%></a> |
                      <%
							  	rsSite2.movenext
							  	if rsSite2.eof then
								exit for
								end if
								next
							%>
                      <%else%>
        暂时无主页
        <%
								end if
								rsSite2.close
								set rsSite2=nothing
							%>
                      </marquee>
                  </div></td>
                </tr>
                <tr >
                  <td height="30" colspan="2"><div align="center"> </div></td>
                </tr>
                <tr>
                  <td height="5" colspan="2">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2"  valign="middle"><div align="center">
                      <marquee behavior="scroll" align="absmiddle" direction="left" width="120" height="25" loop="-1" scrollamount="80" scrolldelay="1000" onMouseOver="javascript: this.stop()" onMouseOut="javascript:this.start()">
                      <%	set rsSite2=server.createobject("adodb.recordset")
								sitesql="select * from user_site order by site_name"
								rsSite2.open sitesql,conn,1,1
							%>
                      <%if not rsSite2.eof then
							for i=0 to rsSite2.recordcount
							%>
                      <a class="style2" href="<%=rsSite2("site_url")%>" title="<%=rsSite2("site_admin")%> : <%=rsSite2("site_info")%>"><%=rsSite2("site_name")%></a> |
                      <%
							  	rsSite2.movenext
							  	if rsSite2.eof then
								exit for
								end if
								next
							%>
                      <%else%>
        暂时无主页
        <%
								end if
								rsSite2.close
								set rsSite2=nothing
							%>
                      </marquee>
                  </div></td>
                </tr>
                <tr valign="bottom">
                  <td width="50%" height="40"><div align="center"><a href="../site_login.asp" class="style3">登陆</a></div></td>
                  <td width="85" height="30"><div align="center"><a href="../site_reg.asp" class="style3">注册</a></div></td>
                </tr>
              </table>
          </div></td>
        </tr>
    </table></td>
  </tr>
</table>
