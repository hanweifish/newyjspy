<%
Set jmail = Server.CreateObject("JMAIL.SMTPMail") '创建一个JMAIL对象
jmail.silent = true 'JMAIL不会抛出例外错误，返回的值为FALSE跟TRUE
jmail.logging = true '启用使用日志
jmail.Charset = "GB2312" '邮件文字的代码为简体中文
jmail.ContentType = "text/html" '邮件的格式为HTML的
jmail.ServerAddress = "smtp.nju.edu.cn" '发送邮件的服务器
jmail.AddRecipient trim(request("address")) '邮件的收件人
jmail.SenderName = trim(request("mail_from")) '邮件发送者的姓名
jmail.Priority = 3
jmail.Subject = trim(request("subject")) '邮件的标题
jmail.Body = trim(request("content")) '邮件的内容
jmail.Execute() '执行邮件发送
jmail.Close '关闭邮件对象
response.write"<script>alert('邮件发送成功！');document.location.href='mail_set.asp';</script>"
%>
