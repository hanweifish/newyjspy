<%
	dim fs,filename,txt,content,total,counter_lenth
	counter_lenth=1  '������ʾ���ݵ���С���ȣ����С��ʵ�ʳ�������ʵ�ʳ���Ϊ׼
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	filename=server.MapPath("count.txt")
	if not fs.FileExists(filename) then
		fs.CreateTextFile filename,True,True
		set txt=fs.OpenTextFile(filename,2,true)
		txt.write 0 '�粻���ڱ������ݵ��ļ��򴴽����ļ���д������0
		set fs=nothing
	end if
	
	set txt=fs.OpenTextFile(filename)
	If txt.AtEndOfStream Then
		Application("Counter")=0 '����ļ���û�����ݣ����ʼ��Application("Counter")��ֵ��Ϊ���ݴ�
	else
		Application("Counter")=txt.readline
	end if

  Application.Lock 
  Application("Counter") = Application("Counter") + 1
  Application.UnLock
  

  Function save_ '�����������
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	filename=server.MapPath("count.txt")
	content=Application("Counter")
	set txt=fs.OpenTextFile(filename,2,true)
	txt.write content
	set fs=nothing
  End Function

  save_  '���ñ��溯����������

  Function Digital ( counter )  '��ʾ���ݺ���
    Dim i,MyStr,sCounter
     sCounter = CStr(counter)
    For i = 1 To counter_lenth - Len(sCounter)
      MyStr = MyStr & "0"
	  'MyStr = MyStr & "<IMG SRC=�ĳ����Լ���ͼƬ��ŵ����Ŀ¼\0.gif>" '����ͼƬ�����ô�������
    Next
    For i = 1 To Len(sCounter)
      MyStr = MyStr & Mid(sCounter, i, 1)
	  'MyStr = MyStr & "<IMG SRC=�ĳ����Լ���ͼƬ��ŵ����Ŀ¼\" & Mid(sCounter, i, 1) & ".gif>" '����ͼƬ�����ô�������
    Next
    Digital = MyStr
  End Function

  Function read_  '��ȡ��������
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	filename=server.MapPath("count.txt")
    set txt=fs.opentextfile(filename,1,true)
	total=txt.readline
	response.write total
	set fs=nothing
  End Function

%>