<%
Class Cls_Setting

	Public Function UpDate()
		Dim Rs,GetValue
		set rs=server.createobject("ADODB.Recordset")
		sql="Select [Title],[Name],[Value],[Data] From ["&tbname&"_Config]"
		rs.open sql,conn,1,3
		Do While Not Rs.Eof
			GetValue = Request.Form("o" & Rs(1) & "")
			Select Case LCase(Rs(3))
			Case "int" : If Len(GetValue) = 0 Or Not IsNumeric(GetValue) Then GetValue = 0
			End Select
			If Rs(1)="Color1" or Rs(1)="Color2" then
				Rs(2) = Replace(GetValue,"#","")
			Elseif Rs(1)="BadWord1" then
				Rs(2) = Replace(Replace(Trim(GetValue),CHR(13),""),CHR(10),"|||")
			Else
				Rs(2) = GetValue
			End if
			Rs(2) = Replace(Rs(2),"%","")
			Rs.Update
			Rs.MoveNext
		Loop
		Rs.Close : Set Rs = Nothing
		Call Refresh()
		UpDate = True
	End Function

	Public Function Refresh()
		call saveconfig() ' ���������ļ�
		Refresh = True
	End Function
End Class

Public function saveconfig()
	dim rs, vupdatepath, config
	vupdatepath = false
	config = config &  "<" & "%"
	config = config & vbcrLf & "Dim dataname"
	config = config & vbcrLf & "    '���ݿ����ƣ����ݿ�λ��DataĿ¼�£��������ƺ���������Ӧ�޸�"
	config = config & vbcrLf &vbcrLf & "    dataname = """&dataname&""""
	config = config & vbcrLf &vbcrLf & "Dim SitePath"
	config = config & vbcrLf & "    '��վ��װĿ¼����/laoy4/,��Ŀ¼��д/����"
	config = config & vbcrLf &vbcrLf & "    SitePath = """&SitePath&""""
	config = config & vbcrLf &vbcrLf & "Dim SiteAdmin"
	config = config & vbcrLf & "    '��̨����Ŀ¼"
	config = config & vbcrLf &vbcrLf & "    SiteAdmin = """&SiteAdmin&""""
	set rs=server.createobject("ADODB.Recordset")
	sql="select [Title],[Name],[value],[Data],[Values] From ["&tbname&"_config] Order By [Order] Asc"
	rs.open sql,conn,1,3
	Do while not rs.eof
	config = config & vbcrLf & vbcrLf & "dim " & lcase(rs(1))
	config = config & vbcrLf & "    '" & lcase(rs(0))
	select case lcase(rs(3))
	case "text"
		config = config & vbcrLf & "    " & lcase(rs(1)) & " = """ & replace(replace(rs(2), """", """"""), vbcrLf, "") & """"
	case "int"
		config = config & vbcrLf & "    " & lcase(rs(1)) & " = " & rs(2)
	end select
	rs.movenext
	Loop
	rs.close: set rs = nothing
	config = config & vbcrLf &vbcrLf & "Dim v"&"e"&"r"&"s"&"i"&"o"&"n"
	config = config & vbcrLf & "    v"&"e"&"r"&"s"&"i"&"o"&"n = ""V"&"4"&"."&"0"&"."&"6"""
	config = config & vbcrLf & "%" & ">"
	Call CreateFile(config,"../inc/config.asp")
end function
%>