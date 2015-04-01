			<DIV id=rolllink 
            style="OVERFLOW: hidden; WIDTH: 130px; HEIGHT: 180px">
            <DIV id=rolllink1>
            <TABLE width="130" border="0" align=center cellpadding="0" cellSpacing=0>
              <TBODY>
								  <tr>
									<td colspan="3" height="20"><div align="center"></div></td>
								  </tr>
								  <tr>
									<td colspan="3" height="20"><div align="center"><font class="style3"> ---- 公 告 ----</font></div></td>
								  </tr>
								<%
									if not (rsb.eof and rsb.bof) then
									for i=1 to rsb.recordcount
								%>
                                      <tr>
                                        <td height="5"></td>
                                        <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                                      </tr>
                                      <tr>
                                        <td width="9%" height="28">&nbsp;</td>
                                        <td width="8%"><div align="center"><img src="indeximages/triangle.gif" width="6" height="8" align="middle"></div></td>
                                        <td width="83%"><font class="style2"><a href=policy_detail.asp?policy_ID=<%=rsb("policy_ID")%> target="_blank" class="style3"><%=rsb("policy_title")%></a></font></td>
                                      </tr>
								<%
									rsb.movenext
									next
								%>
								<%
									else
								%>
								  <tr>
									<td colspan="3"><div align="center"><font class="style3" >暂时没有新的政策发布！</font></div></td>
								  </tr>
								<%
									end if
								%>
								<%
									rsb.close
									set rsb=nothing
								%>
              </TBODY>
			  </TABLE>
            </DIV>
            <DIV id=rolllink2></DIV></DIV>

<script>
   var rollspeed=30
   rolllink2.innerHTML=rolllink1.innerHTML //克隆rolllink1为rolllink2
   function Marquee(){
   if(rolllink2.offsetTop-rolllink.scrollTop<=0) //当滚动至rolllink1与rolllink2交界时
   rolllink.scrollTop-=rolllink1.offsetHeight  //rolllink跳到最顶端
   else{
   rolllink.scrollTop++
   }
   }
   var MyMar=setInterval(Marquee,rollspeed) //设置定时器
   rolllink.onmouseover=function() {clearInterval(MyMar)}//鼠标移上时清除定时器达到滚动停止的目的
   rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}//鼠标移开时重设定时器
</SCRIPT>
