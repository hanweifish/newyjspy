<%
Set jmail = Server.CreateObject("JMAIL.SMTPMail") '����һ��JMAIL����
jmail.silent = true 'JMAIL�����׳�������󣬷��ص�ֵΪFALSE��TRUE
jmail.logging = true '����ʹ����־
jmail.Charset = "GB2312" '�ʼ����ֵĴ���Ϊ��������
jmail.ContentType = "text/html" '�ʼ��ĸ�ʽΪHTML��
jmail.ServerAddress = "smtp.nju.edu.cn" '�����ʼ��ķ�����
jmail.AddRecipient trim(request("address")) '�ʼ����ռ���
jmail.SenderName = trim(request("mail_from")) '�ʼ������ߵ�����
jmail.Priority = 3
jmail.Subject = trim(request("subject")) '�ʼ��ı���
jmail.Body = trim(request("content")) '�ʼ�������
jmail.Execute() 'ִ���ʼ�����
jmail.Close '�ر��ʼ�����
response.write"<script>alert('�ʼ����ͳɹ���');document.location.href='mail_set.asp';</script>"
%>
