<%dim today
today=Date 
today=Year(today) & "-" & Right("0" & Month(today),2) & "-" & Right("0" & Day(today),2)
%>
<%
if left(time(),2) = "ÏÂÎç" then
today = today&" "&CStr(CInt(left(right(time(),8),2))+12)&right(time(),6)
else
today = today&" "&right(time(),8)
end if
%>
