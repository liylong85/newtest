<%
Class Cls_Template
	Dim Reg
	Dim Content
	Dim Template
	
	Private Sub Class_Initialize()
		Set Reg = New RegExp
		Reg.Ignorecase = True
		Reg.Global = True
		Content = ""
		Template = "" ' ģ��·��
	End Sub
	
	Public Function Load(ByVal Templatefile)
		Template = Templatefile
		Call Loadfile
	End Function

	Public Function Parser_Sys()
		'On Error Resume Next
		Dim Matches, Match, SysValue
		Reg.Pattern = "{laoy_([\s\S]*?)}"
		Set Matches = Reg.Execute(Content)
		For Each Match In Matches
			If InStr(LCase(Match.SubMatches(0)), "database") = 0 Then
				If Len(Replace(Match.SubMatches(0), " ", "")) > 0 Then Execute ("SysValue = " & Replace(Match.SubMatches(0), " ", "")) Else SysValue = ""
			Else
				SysValue = ""
			End If
			Content = Replace(Content, Match.Value, SysValue) ' �滻
			If Err Then Err.Clear: Echo "<font color=red>ϵͳ��ǩ����</font>": Response.End
		Next
		Content=Replace(Content,"{$laoy_��վͷ��}",Head)
		Content=Replace(Content,"{$laoy_��վ����}",Menu)
		Content=Replace(Content,"{$laoy_��վβ��}",Copy1)	
	End Function
		
	' ����ģ��
	Public Function Loadfile()
		Dim Obj
		'On Error Resume Next
		Set Obj = Server.CreateObject("adodb.stream")
		With Obj
		.Type = 2: .Mode = 3: .Open: .Charset = "GB2312" : .Position = Obj.Size: .Loadfromfile Server.Mappath(Template): Content = .ReadText: .Close
		End With
		Set Obj = Nothing
		If Err Then Echo "<font color=red>" & Err.Description & "</font>": Response.End
	End Function
End Class
%>