<!--#include file="../Inc/conn.asp"-->
<!--#include file="Admin_check.asp"-->
<%
If yaoadmintype<>1 then
Response.Redirect ""&SitePath&""&SiteAdmin&"/Info.asp?Info=û��Ȩ��"
Response.End()
End If
%><html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equtv="Content-Type" content="text/html; charset=gb2312" />
<title>��վ��̨����--�ļ�����</title>
<link href="images/Admin_css.css" type=text/css rel=stylesheet>
</head>
<body>
<%	
	if request("action")="del" then
		call DelOver()
	else
		call Main()
	end if
Sub Main
%>
<table border="0"  cellspacing="1" cellpadding="3" height="1" class="admintable1">
<tr>
<td class="admintitle">�ϴ��ļ�����</td>
</tr><tr>
<td height=30 bgcolor="#FFFFFF" class="td"><br><br><b>ע�⣺</b>�����ʹ�ñ����ܣ�����ǰ����ȱ������ϴ��ļ��У������ļ���ʧ��<br><br>
�����������ر��(����)�������������ܻ�ǳ��������Ƽ�ʹ�ñ����ܡ�
<br><br>��ǰ�ϴ��ļ���Ϊ��<%=SiteUp%>
  <br /><br /><input type="button" value="��ʼ����" class="bnt" onClick="if(confirm('ȷ��Ҫ����ļ���<%=SiteUp%>�е������ļ�?\n\n�����������'))location.href='?action=del';return false;" />
  <br /><br /><br />
</td>
</tr>
</table>
  <%
End Sub

Sub DelOver
	Dim Sql,Rs,strFiles,ItemIntro
	strFiles="|"&SitePath&SiteUp&"|"&SiteLogo
	ItemIntro=""
	
	Sql = "Select Content,images From "&tbname&"_Article"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End if
		IF Rs(1) <> "" Then
			strFiles=strFiles&"|"&Rs(1)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	
	Sql = "Select Content From "&tbname&"_Ad"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	
	Sql = "Select LogoUrl From "&tbname&"_Link"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|"&Rs(0)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	
	Sql = "Select UserFace From "&tbname&"_User"
	Set Rs = Conn.Execute(Sql)
	While Not Rs.Eof
		IF Rs(0) <> "" Then
			strFiles=strFiles&"|UserFace/"&Rs(0)
		End If
	Rs.MoveNext
	Wend
	Rs.Close
	
	Dim tempStr, tempi, TempArray,UpFileType,regEx,Matches,Match
		UpFileType = "gif|jpg|jpeg|bmp|png"
		Set regEx=New Regexp
		regEx.Ignorecase=True
		regEx.Global=True
        regEx.Pattern = "<img.+?[^\>]>" '��ѯ���������� <img..>
        Set Matches = regEx.Execute(ItemIntro)
        For Each Match In Matches
            If tempStr <> "" Then
                tempStr = tempStr & "|" & Match.value '�ۼ�����
            Else
                tempStr = Match.value
            End If
        Next
        If tempStr <> "" Then
            TempArray = Split(tempStr, "|") '�ָ�����
            tempStr = ""
            For tempi = 0 To UBound(TempArray)
                regEx.Pattern = "src\s*=\s*.+?\.(" & UpFileType & ")" '��ѯsrc =�ڵ�����
                Set Matches = regEx.Execute(TempArray(tempi))
                For Each Match In Matches
                    If tempStr <> "" Then
                        tempStr = tempStr & "|" & Match.value '�ۼӵõ� ���Ӽ�$Array$ �ַ�
                    Else
                        tempStr = Match.value
                    End If
                Next
            Next
        End If
        If tempStr <> "" Then
            regEx.Pattern = "src\s*=\s*" '���� src =
            tempStr = regEx.Replace(tempStr, "")
        End If
		Set regEx=Nothing
        strFiles = strFiles & tempStr
	    strFiles = LCase(strFiles)
		Dim i,theFolder,fso,theFile,theSubFolder
		Set Fso=CreateObject("Scripting.FileSystemObject")
		 i = 0
     

    Set theFolder = fso.GetFolder(Server.MapPath(SitePath&SiteUp))
    For Each theFile In theFolder.Files
        If InStr(strFiles, LCase(theFile.name)) <= 0 Then
            theFile.Delete True
            i = i + 1
        End If
    Next
    For Each theSubFolder In theFolder.SubFolders
        For Each theFile In theSubFolder.Files
            If InStr(strFiles, LCase(theSubFolder.name & "/" & theFile.name)) <= 0 Then
                theFile.Delete True
                i = i + 1
            End If
        Next
    Next
	Call Alert("���ι�����"&I&"�������ļ�","?")
End Sub
%>
<!--#include file="Admin_copy.asp"-->
</body>
</html>