<%
	strSourceFile = Server.MapPath("../../inc/config.xml")
	Set objXML = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")
	objXML.load(strSourceFile)
	Set objRoot = objXML.selectSingleNode("Config")
	dim info
	info=objRoot.childNodes.item(0).text
%>
<table width="603" height="236" border="0" cellpadding="0" cellspacing="0" background="../../includeimages/service.gif">
  <tr>
    <td width="406" height="60" valign="middle"><div align="center"><img src="../../includeimages/phone.gif" width="150" height="35"></div></td>
    <td width="3" rowspan="3"><img src="../../includeimages/bar1.gif" width="3" height="146"></td>
    <td width="192" rowspan="3" valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="200" height="60" valign="middle"><div align="center"><a href="../../Site_login.asp"><img src="../../includeimages/stusite.gif" width="150" height="35" border="0"></a></div></td>
        </tr>
        <tr>
          <td width="200" height="176" valign="top"><div align="center">
              <table width="150" height="60%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td colspan="2" bgcolor="#DCE4F3"><div align="center">
                      <marquee behavior="scroll" direction="right" width="160" height="30" loop="-1" scrollamount="80" scrolldelay="1000" onMouseOver="javascript: this.stop()" onMouseOut="javascript:this.start()">
                      <table height="100%"  border="0" cellpadding="0" cellspacing="0">
                        <%	set rsSite2=server.createobject("adodb.recordset")
								sitesql="select * from user_site order by site_name"
								rsSite2.open sitesql,conn,1,1
							%>
                        <tr valign="middle">
                          <%if not rsSite2.eof then
							for i=0 to rsSite2.recordcount
							%>
                          <td width="80" height="35"><div align="center"><a class="style2" href="<%=rsSite2("site_url")%>" title="<%=rsSite2("site_admin")%> : <%=rsSite2("site_info")%>"><%=rsSite2("site_name")%></a></div></td>
                          <%
							  	rsSite2.movenext
							  	if rsSite2.eof then
								exit for
								end if
								next
							%>
                          <%else%>
                          <td width="100" class="style2"><div align="center">&#26242;&#26102;&#26080;&#20027;&#39029;</div></td>
                          <%
								end if
								rsSite2.close
								set rsSite2=nothing
							%>
                        </tr>
                      </table>
                      </marquee>
                  </div></td>
                </tr>
                <tr bgcolor="#FFFFEE">
                  <td height="35" colspan="2"><div align="center"> </div></td>
                </tr>
                <tr>
                  <td colspan="2" bgcolor="#DCE4F3"><div align="center">
                      <marquee behavior="scroll" direction="left" width="160" height="30" loop="-1" scrollamount="80" scrolldelay="1000" onMouseOver="javascript: this.stop()" onMouseOut="javascript:this.start()">
                      <table height="100%"  border="0" cellpadding="0" cellspacing="0">
                        <%	set rsSite2=server.createobject("adodb.recordset")
								sitesql="select * from user_site order by site_name"
								rsSite2.open sitesql,conn,1,1
							%>
                        <tr valign="middle">
                          <%if not rsSite2.eof then
							for i=0 to rsSite2.recordcount
							%>
                          <td width="80" height="35"><div align="center"><a class="style2" href="<%=rsSite2("site_url")%>" title="<%=rsSite2("site_admin")%> : <%=rsSite2("site_info")%>"><%=rsSite2("site_name")%></a></div></td>
                          <%
							  	rsSite2.movenext
							  	if rsSite2.eof then
								exit for
								end if
								next
							%>
                          <%else%>
                          <td width="100" class="style2"><div align="center">&#26242;&#26102;&#26080;&#20027;&#39029;</div></td>
                          <%
								end if
								rsSite2.close
								set rsSite2=nothing
							%>
                        </tr>
                      </table>
                      </marquee>
                  </div></td>
                </tr>
                <tr>
                  <td height="30" colspan="2"><div align="right"></div></td>
                </tr>
                <tr>
                  <td width="50%" height="20"><div align="center"><a href="../../site_login.asp" class="style3">&#30331;&#38470;</a></div></td>
                  <td width="85"><div align="center"><a href="../../site_reg.asp" class="style3">&#27880;&#20876;</a></div></td>
                </tr>
              </table>
          </div></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><div valign="top" align="center">
          <table width="90%" height="90%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#D6E0F1">
            <tr>
              <td width="54%" height="100" class="style2"><div align="left">&nbsp;&nbsp;&nbsp;&nbsp;<%=info%></div>                <div align="left"></div>                </td>
            </tr>
          </table>
    </div>      <div valign="top" align="center">
        <div align="left">        </div>
      </div></td>
  </tr>
  <tr>
    <td valign="middle">&nbsp;</td>
  </tr>
</table>
