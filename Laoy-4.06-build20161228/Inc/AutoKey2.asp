<%
Function cn_split(str, x)
    Dim a, b 'As String
    Dim i, j, flag, max, temp_str

	Dim Fso,MyFso,TxtContent
	Set Fso=Server.CreateObject("Scripting.FileSystemObject")
	Set MyFso=Fso.OpenTextFile(Server.mappath(SitePath &"data/zwck.txt"))
	TxtContent=MyFso.ReadAll
	b="|" & Replace(TxtContent,vbcrlf,"|") & "|"
	Set MyFso=nothing
	Set Fso=nothing

    a = "laoy"&str

    '�ִ�
    For i = 1 To Len(a)
        For j = 2 To x
            a = a & Mid(a, i, j) & " "
        Next
    Next

    a = Split(a, " ")
    'b = Split(b, " ")
    max = UBound(a)
    '�����ظ��ַ���
    For i = 0 To max - 1
			flag = a(i)
			'If iscn(flag) Then
				For j = i + 1 To max - 1
					If a(j) = flag And flag <> "" Then
						a(j) = ""   
					End If
				Next
				If a(i) <> "" Then
					temp_str = temp_str & a(i) & " "
				End If
		   'End If
	Next

    	a = Split(temp_str, " ")
		temp_str = ""
		 For i=0 to Ubound(a)
		  If Instr(b,"|" & a(i) & "|")<>0 Then
		   temp_str = temp_str  &"<a href=""../bbs/?KeyWord="&Server.UrlEncode(a(i))&""">"& a(i) &"</a>"&  ","
		   'If len(temp_str)>maxlen and maxlen<>0 then exit for
		  end if
		 Next
		if temp_str<>"" then
		cn_split="<div class=""bbskey"">�ؼ��֣�"&replace(trim(temp_str),","," ")&"</div>"
		else
		cn_split=""
		end if
End Function
	
	function iscn(str) 
		Dim i 
		i = Len(str) 
		If i = 0 Then 
		   iscn = False 
		   Exit Function 
		End If 
		
		Do While i > 0 
		  If Asc(Mid(str, i, 1)) < 10000 And Asc(Mid(str, i, 1)) > -10000 Then 
		    iscn = False 
		    Exit Function 
		  End If 
		  i = i - 1 
		Loop 
		iscn = True 
	end function
%>