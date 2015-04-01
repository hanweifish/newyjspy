<table width="603" height="236" border="0" cellpadding="0" cellspacing="0" background="includeimages/service.gif">
  <tr>
    <td width="406" colspan="2" valign="middle"><div align="center"></div>      <div valign="top" align="center">
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="60" valign="middle"><div align="center"><font color="#FF6600">&#26368;&#26032;出国信息</font></div></td>
        </tr>
        <tr>
          <td height="176" valign="top"><div align="center">
              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><div align="center">
                    <marquee behavior="scroll" direction="up"  width="300" height="120" loop="-1" scrollamount="1" scrolldelay="1" >
                    <%	
										set rsnews=server.createobject("adodb.recordset")
										newsql="select * from news order by news.news_time desc"
										rsnews.open newsql,conn,1,1
									%>
                    <table width="300"  border="0" cellpadding="0" cellspacing="0">
					<%
										if not rsnews.eof then
										for i=0 to rsnews.recordcount
									%>
                      <tr>
                        <td valign="middle"><div align="center"><img src="indeximages/triangle.gif" width="6" height="8"></div></td>
                        <td width="285" height="25" valign="middle"><div align="left"><a href="news_detail.asp?news_id=<%=rsnews("news_ID")%>" target="_blank"><%=rsnews("news_title")%></a></div></td>
                      </tr>
					  <%
							rsnews.movenext
							if rsnews.eof then
							exit for
							end if
							next
					  %>
                    </table>
                    <%
						else
					%>
                    <span class="style3">&#27809;&#26377;&#26032;&#38395;&#21457;&#24067; </span>
                    <%
							end if
							rsnews.close
							set rsnews=nothing
						%>
                    </marquee>
                  </div></td>
                </tr>
                <tr>
                  <td height="30" valign="bottom"><div align="center"><a href="newslist.asp" class="style3" target="_blank">&lt;&lt; MORE &gt;&gt; </a></div></td>
                </tr>
              </table>
          </div></td>
        </tr>
      </table>
    </div>      </td>
    <td width="3"><img src="includeimages/bar1.gif" width="3" height="146"></td>
    <td width="192" valign="top"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="200" height="60" valign="middle"><div align="center">个人主页推荐</div></td>
        </tr>
        <tr>
          <td width="200" height="176" valign="top"><div align="center">
              <table width="150" height="60%"  border="0" cellpadding="0" cellspacing="0">
                <tr><td height="5" colspan="2" >&nbsp;</td>
                </tr>
				<tr>
                  <td colspan="2"  valign="middle"><div align="center">
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
                <tr>
                  <td height="30" colspan="2"><div align="center"> </div></td>
                </tr>
                <tr >
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
                  <td width="50%" height="40"><div align="center"><a href="site_login.asp" class="style3">登陆</a></div></td>
                  <td width="85" height="30"><div align="center"><a href="site_reg.asp" class="style3">注册</a></div></td>
                </tr>
              </table>
          </div></td>
        </tr>
    </table></td>
  </tr>
</table>
