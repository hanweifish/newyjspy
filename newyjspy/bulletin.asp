			<DIV id=rolllink 
            style="OVERFLOW: hidden; WIDTH: 130px; HEIGHT: 180px">
            <DIV id=rolllink1>
            <TABLE width="130" border="0" align=center cellpadding="0" cellSpacing=0>
              <TBODY>
								  <tr>
									<td colspan="3" height="20"><div align="center"></div></td>
								  </tr>
								  <tr>
									<td colspan="3" height="20"><div align="center"><font class="style3"> ---- �� �� ----</font></div></td>
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
									<td colspan="3"><div align="center"><font class="style3" >��ʱû���µ����߷�����</font></div></td>
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
   rolllink2.innerHTML=rolllink1.innerHTML //��¡rolllink1Ϊrolllink2
   function Marquee(){
   if(rolllink2.offsetTop-rolllink.scrollTop<=0) //��������rolllink1��rolllink2����ʱ
   rolllink.scrollTop-=rolllink1.offsetHeight  //rolllink�������
   else{
   rolllink.scrollTop++
   }
   }
   var MyMar=setInterval(Marquee,rollspeed) //���ö�ʱ��
   rolllink.onmouseover=function() {clearInterval(MyMar)}//�������ʱ�����ʱ���ﵽ����ֹͣ��Ŀ��
   rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}//����ƿ�ʱ���趨ʱ��
</SCRIPT>
