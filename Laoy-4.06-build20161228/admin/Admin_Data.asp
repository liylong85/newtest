<!--#include file="../Inc/conn.asp"-->
<!--#include file="Admin_check.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��վ��̨����</title>
<LINK href="images/Admin_css.css" type=text/css rel=stylesheet>
<script src="js/admin.js"></script>
</head>

<body>
<%
select case request("action")
    case "SpaceSize"
		Call chkAdmin(17)
	    SpaceSize()
	case "CompressData"
		Call chkAdmin(14)
		if IsSqlDataBase = 1 then
			SQLUserReadme()
		else
			CompressData()
		end if
	case "BackupData"
		Call chkAdmin(15)
	    if request("act")="Backup" then
			if IsSqlDataBase = 1 then
				SQLUserReadme()
			else
				call updata()
			end if
		else
			if IsSqlDataBase = 1 then
				SQLUserReadme()
			else
				call BackupData()
			end if
		end if
	case "RestoreData"
		Call chkAdmin(16)
		if IsSqlDataBase = 1 then
			SQLUserReadme()
		else
			if request("act")="Restore" then
				dim Dbpath,backpath
				Dbpath=request.form("Dbpath")
				backpath=request.form("backpath")
				if dbpath="" then
				Call Alert("��������Ҫ�ָ������ݿ�·��������!","-1")	
				else
				Dbpath=server.mappath(Dbpath)
				end if
				backpath=server.mappath(backpath)
			
				Set Fso=Server.CreateObject("Scripting.FileSystemObject")
				if fso.fileexists(dbpath) then  					
					fso.copyfile Dbpath,Backpath
					Call Alert("�ɹ��ָ�����!","Admin_data.asp?action=SpaceSize")	
				else
					Call Alert("����Ŀ¼�²������ı����ļ�!","-1")	
				end if
			else
				call RestoreData()
			end if
		End if
end select
%>
<%
'====================ѹ�����ݿ� =========================
sub CompressData()
%>
<table border="0"  cellspacing="1" cellpadding="3" height="1" class="admintable1">
<form action="Admin_data.asp?action=CompressData" method="post">
<tr>
<td class="admintitle">ѹ�����ݿ�</td>
</tr><tr>
<td height=30 bgcolor="#FFFFFF" class="td"><b>ע�⣺</b>�������ݿ��������·��,�����������ݿ����ƣ�����ʹ�������ݿ���ܻ�ѹ��ʧ�ܣ���ѡ�񱸷����ݿ����ѹ�������� </td>
</tr>
<tr>
<td height="30" bgcolor="#FFFFFF" class="td">ѹ�����ݿ⣺<input name="dbpath" type="text" value="../Data/<%=DataName%>" size="50">
&nbsp;
<input type="submit" class="bnt" value="��ʼѹ��"></td>
</tr>
<tr>
<td height="30" bgcolor="#FFFFFF" class="td"><input name="boolIs97" type="checkbox" class="noborder" value="True">���ʹ�� Access 97 ���ݿ���ѡ��
(Ĭ��Ϊ Access 2000 ���ݿ�)<br></td>
</tr>
<tr>
  <td height="30" bgcolor="#FFFFFF" class="td">ע���뾡����ftp���ػ����ݿ��ѹ�������������������Ҫʹ�ô˹��ܣ��뱸�ݺ���ѹ����<strong>���ݿ�������������߸Ų�����!</strong></td>
</tr>
</form>
</table>
<%
dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If

end sub

'=====================ѹ������=========================
Function CompactDB(dbPath, boolIs97)
	Dim fso, Engine, strDBPath,JET_3X
	strDBPath = left(dbPath,instrrev(DBPath,"\"))
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	If fso.FileExists(dbPath) Then
		fso.CopyFile dbpath,strDBPath & "temp.mdb"
		Set Engine = CreateObject("JRO.JetEngine")
	
		If boolIs97 = "True" Then
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" _
			& "Jet OLEDB:Engine Type=" & JET_3X
		Else
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
		End If
	
		fso.CopyFile strDBPath & "temp1.mdb",dbpath
		fso.DeleteFile(strDBPath & "temp.mdb")
		fso.DeleteFile(strDBPath & "temp1.mdb")
		Set fso = nothing
		Set Engine = nothing
		
		Call Alert("ѹ���ɹ�!","Admin_data.asp?action=SpaceSize")
	Else
		Call Alert("���ݿ����ƻ�·������ȷ. ������!","-1")
	End If
End Function
%>
<%
'====================�������ݿ�=========================
sub BackupData()
%>
	<table border="0"  cellspacing="1" cellpadding="3" class="admintable1">
	  <tr>
		  <td colspan="2" class="admintitle" >������վϵͳ����( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )</td>
	  	</tr>
  				<form method="post" action="Admin_data.asp?action=BackupData&act=Backup">
  				<tr>
  				  <td width="19%" height="30" align="center" bgcolor="#FFFFFF" class="td">��ǰ���ݿ�·��(���)��</td>
				  <td width="81%" bgcolor="#FFFFFF" class="td"><input name=DBpath type=text id="DBpath" value="../Data/<%=DataName%>" size="40"  Readonly="true"/></td>
  				</tr>
  				<tr>
  				  <td height="30" align="center" bgcolor="#FFFFFF" class="td">�������ݿ�Ŀ¼(���)��</td>
				  <td bgcolor="#FFFFFF" class="td"><input name=bkfolder type=text value="../Databackup/" size="40" Readonly="true"/>
&nbsp;��Ŀ¼�����ڣ������Զ�����</td>
  				</tr>
  				<tr>
  				  <td height="30" align="center" bgcolor="#FFFFFF" class="td">�������ݿ�����(����)��</td>
				  <td bgcolor="#FFFFFF" class="td"><input name=bkDBname type=text value="<%=year(Now)&month(Now)&day(Now)&"_"&hour(Now)&Minute(Now)&Second(Now)%>.mdb" size="40"  Readonly="true"/>
&nbsp;�籸��Ŀ¼�и��ļ��������ǣ���û�У����Զ�����</td>
  				</tr>
  				<tr>
  				  <td height="30" bgcolor="#FFFFFF" class="td">&nbsp;</td>
				  <td bgcolor="#FFFFFF" class="td"><input type=submit class="bnt" value="ȷ��" /></td>
  				</tr>
  				<tr>
  				  <td height="30" bgcolor="#FFFFFF" class="td">&nbsp;</td>
				  <td bgcolor="#FFFFFF" class="td">ע�⣺����·��������������ռ��Ŀ¼�����·�� </td>
  				</tr>	
  				</form>
  </table>
<%
end sub

sub updata()
	dim Dbpath,bkfolder,bkdbname,fso
	Dbpath="../Data/"&DataName&""
	Dbpath=server.mappath(Dbpath)
	bkfolder=request.form("bkfolder")
	bkdbname=request.form("bkdbname")
	Set Fso=Server.CreateObject("Scripting.FileSystemObject")
	if fso.fileexists(dbpath) then
		If CheckDir(bkfolder) = True Then
		fso.copyfile dbpath,bkfolder& "\"& bkdbname
		else
		MakeNewsDir bkfolder
		fso.copyfile dbpath,bkfolder& "\"& bkdbname
		end if		
		Call Alert ("�������ݿ�ɹ�!","Admin_data.asp?action=SpaceSize")
	Else
		Call Alert ("���ݿ�·������","-1")
	End if
end sub
'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
    dim fso1
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '����
       CheckDir = True
    Else
       '������
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------����ָ����������Ŀ¼-----------------------
Function MakeNewsDir(foldername)
	dim f,fso1
    Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function
%>
<%
'====================�ָ����ݿ�=========================
sub RestoreData()
%>
<table border="0"  cellspacing="1" cellpadding="3" class="admintable1">
	<tr>
		<td colspan="2" class="admintitle">�ָ���վϵͳ����( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )</td>
    </tr>
<form method="post" action="Admin_data.asp?action=RestoreData&act=Restore">
<tr>
  <td width="19%" height="30" align="center" bgcolor="#FFFFFF" class="td">�������ݿ�·��(���)��</td>
  <td width="81%" bgcolor="#FFFFFF" class="td"><select name="DBpath" id="DBpath">
  <%
  Dim bflist
  bflist=FileList("../Databackup","mdb")
  If bflist<>"" then
  bflist1=split(left(bflist,len(bflist)-1),"|")
  For i=0 to Ubound(bflist1)
  %>
    <option value="../DataBackup/<%=bflist1(i)%>"><%=bflist1(i)%></option><%
	Next
	Else
	%><option value="">���ź���û�б������ݿ�</option><%
	End if
	%>
  </select>
  </td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#FFFFFF" class="td">Ŀ�����ݿ�·��(���)��</td>
  <td bgcolor="#FFFFFF" class="td"><input type=text size=40 name=backpath value="../Data/<%=DataName%>"  Readonly="true"/></td>
</tr>
<tr>
  <td height="30" bgcolor="#FFFFFF" class="td">&nbsp;</td>
  <td bgcolor="#FFFFFF" class="td"><input type=submit class="bnt" value="�ָ�����" /></td>
</tr>
<tr>
  <td height="30" bgcolor="#FFFFFF" class="td">&nbsp;</td>
  <td bgcolor="#FFFFFF" class="td">�������������������Ų�����</td>
</tr>	
</form>
</table>
<%
end sub
sub SpaceSize()
%>
<table height="1" border="0" cellpadding="3"  cellspacing="1" bgcolor="#F2F9E8" class="admintable1">
  <tr>
    <td colspan="2" class="admintitle">����ռ�ÿռ���� </td>
  </tr>
  <tr>
    <td width="19%" height="30" align="center" bgcolor="#FFFFFF">ϵͳռ�ÿռ��ܼƣ�</td>
    <td width="81%" bgcolor="#FFFFFF"><img src="images/bar1.gif" width=400 height=9 />&nbsp;
    <%allsize()%></td>
  </tr>
  <tr>
    <td height="30" align="center" bgcolor="#FFFFFF">��������ռ�ÿռ䣺</td>
    <td bgcolor="#FFFFFF"><img src="images/bar1.gif" width=<%=drawbar("Data")%> height=9 />&nbsp;<%othersize("Data")%></td>
  </tr>
  <tr>
    <td height="30" align="center" bgcolor="#FFFFFF">��������ռ�ÿռ䣺</td>
    <td bgcolor="#FFFFFF"><img src="images/bar1.gif" width=<%=drawbar("Databackup")%> height=9 />&nbsp;<%othersize("Databackup")%></td>
  </tr>
  <tr>
    <td height="30" align="center" bgcolor="#FFFFFF">ϵͳͼƬռ�ÿռ䣺</td>
    <td bgcolor="#FFFFFF"><img src="images/bar1.gif" width=<%=drawbar("Images")%> height=9 />&nbsp;<%othersize("Images")%></td>
  </tr>
  <tr>
    <td height="30" align="center" bgcolor="#FFFFFF">�ϴ�ͼƬռ�ÿռ䣺</td>
    <td bgcolor="#FFFFFF"><img src="images/bar1.gif" width=<%=drawbar(""&SiteUp&"")%> height=9 />&nbsp;<%othersize(SiteUp)%></td>
  </tr>
  <tr>
    <td height="30" align="center" bgcolor="#FFFFFF">��|-�û��ϴ�ͷ��</td>
    <td bgcolor="#FFFFFF"><img src="images/bar1.gif" width=<%=drawbar(""&SiteUp&"/UserFace")%> height=9 />&nbsp;<%othersize(""&SiteUp&"/UserFace")%></td>
  </tr>
</table>
<%end sub%>
<%
sub othersize(names)
	dim fso,path,ml,mlsize,dx,d,size
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	path=server.mappath("..\Images")
	ml=left(path,(instrrev(path,"\")-1))&"\"&names
	
	On Error Resume Next
	set d=fso.getfolder(ml) 
	If Err Then
		err.Clear
		Response.Write "<font color=red>��ʾ��û�С�"&names&"��Ŀ¼</font>"					
		'Response.End()
	End If
	mlsize=d.size
	size=mlsize
	dx=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   dx=formatnumber(size,2) & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   dx=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   dx=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write dx
end sub

sub allsize()
	dim fso,path,ml,mlsize,dx,d,size
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	path=server.mappath("../index.asp")
	ml=left(path,(instrrev(path,"\")-1))
	set d=fso.getfolder(ml) 
	mlsize=d.size
	size=mlsize
	dx=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   dx=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   dx=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   dx=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write dx
end sub

Function Drawbar(drvpath)
	dim fso,drvpathroot,d,size,totalsize,barsize
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	drvpathroot=server.mappath("../Images")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	
	On Error Resume Next
	drvpath=server.mappath("../"&drvpath)
	If Err Then
		err.Clear
		Response.Write "û����Ϊ��"&drvpath&"����Ŀ¼���������޸��ļ�����ȷ��ʾ��Ŀ¼��ʹ������"			
		Response.End()
	End If
	set d=fso.getfolder(drvpath)
	size=d.size
	
	barsize=cint((size/totalsize)*400)
	Drawbar=barsize
End Function 

Function FileList(FolderUrl,FileExName)
Set fso=Server.CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set folder=fso.GetFolder(Server.MapPath(Trim(FolderUrl)))
Set file=folder.Files
FileList=""
For Each FileName in file
If Trim(FileExName)<>"" Then
	If InStr(Trim(FileExName),Trim(Mid(FileName.Name,InStr(FileName.Name,".")+1,len(FileName.Name))))>0 Then
    	FileList=FileList&""&FileName.Name&"|"
	End If
Else
    FileList=FileList&"<a href='#'>"&FileName.Name&"</a><br>"
End If
Next
Set file=Nothing
Set folder=Nothing
Set fso=Nothing
End Function

Sub SQLUserReadme()
'�˲���ժ�Զ�����̳
%>
        <table border="0"  cellspacing="1" cellpadding="3" class="admintable1">
          <tr>
              <td colspan="2" class="admintitle" >SQL���ݿ����ݴ���˵��</td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF" style="font-size:14px;line-height:180%">
            <div style="font:bold 14px Microsoft Yahei,sans-serif;color:#f00;">�����ֻ���õ�������������<u>��ѯ�ռ�����α��ݼ��ָ�SQl���ݿ⡣</u><br>
            ������õĶ�������������VPS����ο��������ݣ�<br></div>
            <blockquote> <B>һ���������ݿ�</B> <BR>
                    <BR>
        1����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server<BR>
        2��SQL Server��-->˫������ķ�����-->˫�������ݿ�Ŀ¼<BR>
        3��ѡ��������ݿ����ƣ�����վ���ݿ�website��-->Ȼ�������˵��еĹ���-->ѡ�񱸷����ݿ�<BR>
        4������ѡ��ѡ����ȫ���ݣ�Ŀ���еı��ݵ����ԭ����·����������ѡ�����Ƶ�ɾ����Ȼ������ӣ����ԭ��û��·����������ֱ��ѡ�����ӣ�����ָ��·�����ļ�����ָ�����ȷ�����ر��ݴ��ڣ����ŵ�ȷ�����б��� <BR>
        <BR>
        <B>������ԭ���ݿ�</B><BR>
        <BR>
        1����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server<BR>
        2��SQL Server��-->˫������ķ�����-->��ͼ�������½����ݿ�ͼ�꣬�½����ݿ����������ȡ<BR>
        3������½��õ����ݿ����ƣ�����վ���ݿ�website��-->Ȼ�������˵��еĹ���-->ѡ��ָ����ݿ�<BR>
        4���ڵ������Ĵ����еĻ�ԭѡ����ѡ����豸-->��ѡ���豸-->������-->Ȼ��ѡ����ı����ļ���-->���Ӻ��ȷ�����أ���ʱ���豸��Ӧ�ó������ղ�ѡ������ݿⱸ���ļ��������ݺ�Ĭ��Ϊ1���������ͬһ���ļ�������α��ݣ����Ե�����ݺ��ԱߵĲ鿴���ݣ��ڸ�ѡ����ѡ�����µ�һ�α��ݺ��ȷ����-->Ȼ�����Ϸ������Աߵ�ѡ�ť<BR>
        5���ڳ��ֵĴ�����ѡ�����������ݿ���ǿ�ƻ�ԭ���Լ��ڻָ����״̬��ѡ��ʹ���ݿ���Լ������е��޷���ԭ����������־��ѡ��ڴ��ڵ��м䲿λ�Ľ����ݿ��ļ���ԭΪ����Ҫ������SQL�İ�װ�������ã�Ҳ����ָ���Լ���Ŀ¼�����߼��ļ�������Ҫ�Ķ������������ļ���Ҫ���������ָ��Ļ���������Ķ���������SQL���ݿ�װ��D:\Program Files\Microsoft SQL Server\MSSQL\Data����ô�Ͱ������ָ�������Ŀ¼������ظĶ��Ķ������������ļ�����øĳ�����ǰ�����ݿ�������ԭ����web_data.mdf�����ڵ����ݿ���website���͸ĳ�website_data.mdf������־�������ļ���Ҫ���������ķ�ʽ����صĸĶ�����־���ļ�����*_log.ldf��β�ģ�������Ļָ�Ŀ¼�������������ã�ǰ���Ǹ�Ŀ¼������ڣ���������ָ��d:\sqldata\web_data.mdf����d:\sqldata\web_log.ldf��������ָ�������<BR>
        6���޸���ɺ󣬵�������ȷ�����лָ�����ʱ�����һ������������ʾ�ָ��Ľ��ȣ��ָ���ɺ�ϵͳ���Զ���ʾ�ɹ������м���ʾ���������¼����صĴ������ݲ�ѯ�ʶ�SQL�����Ƚ���Ϥ����Ա��һ��Ĵ����޷���Ŀ¼��������ļ����ظ������ļ���������߿ռ䲻���������ݿ�����ʹ���еĴ������ݿ�����ʹ�õĴ��������Գ��Թر����й���SQL����Ȼ�����´򿪽��лָ��������������ʾ����ʹ�õĴ�����Խ�SQL����ֹͣȻ�����𿴿����������������Ĵ���һ�㶼�ܰ��մ�����������Ӧ�Ķ��󼴿ɻָ�<BR>
        <BR>
        <B>�����������ݿ�</B><BR>
        <BR>
        һ������£�SQL���ݿ�����������ܴܺ�̶��ϼ�С���ݿ��С������Ҫ������������־��С��Ӧ�����ڽ��д˲����������ݿ���־����<BR>
        1���������ݿ�ģʽΪ��ģʽ����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����-->˫�������ݿ�Ŀ¼-->ѡ��������ݿ����ƣ�����վ���ݿ�website��-->Ȼ�����Ҽ�ѡ������-->ѡ��ѡ��-->�ڹ��ϻ�ԭ��ģʽ��ѡ�񡰼򵥡���Ȼ��ȷ������<BR>
        2���ڵ�ǰ���ݿ��ϵ��Ҽ��������������е��������ݿ⣬һ�������Ĭ�����ò��õ�����ֱ�ӵ�ȷ��<BR>
        3��<font color=blue>�������ݿ���ɺ󣬽��齫�������ݿ�������������Ϊ��׼ģʽ����������ͬ��һ�㣬��Ϊ��־��һЩ�쳣����������ǻָ����ݿ����Ҫ����</font> <BR>
        <BR>
        <B>�ġ��趨ÿ���Զ��������ݿ�</B><BR>
        <BR>
        <font color=red>ǿ�ҽ������������û����д˲�����</font><BR>
        1������ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����<BR>
        2��Ȼ�������˵��еĹ���-->ѡ�����ݿ�ά���ƻ���<BR>
        3����һ��ѡ��Ҫ�����Զ����ݵ�����-->��һ�����������Ż���Ϣ������һ�㲻����ѡ��-->��һ��������������ԣ�Ҳһ�㲻ѡ��<BR>
        4����һ��ָ�����ݿ�ά���ƻ���Ĭ�ϵ���1�ܱ���һ�Σ��������ѡ��ÿ�챸�ݺ��ȷ��<BR>
        5����һ��ָ�����ݵĴ���Ŀ¼��ѡ��ָ��Ŀ¼������������D���½�һ��Ŀ¼�磺d:\databak��Ȼ��������ѡ��ʹ�ô�Ŀ¼������������ݿ�Ƚ϶����ѡ��Ϊÿ�����ݿ⽨����Ŀ¼��Ȼ��ѡ��ɾ�����ڶ�����ǰ�ı��ݣ�һ���趨4��7�죬�⿴���ľ��屸��Ҫ�󣬱����ļ���չ��һ�㶼��bak����Ĭ�ϵ�<BR>
        6����һ��ָ��������־���ݼƻ�����������Ҫ��ѡ��-->��һ��Ҫ���ɵı�����һ�㲻��ѡ��-->��һ��ά���ƻ���ʷ��¼�������Ĭ�ϵ�ѡ��-->��һ�����<BR>
        7����ɺ�ϵͳ�ܿ��ܻ���ʾSql Server Agent����δ�������ȵ�ȷ����ɼƻ��趨��Ȼ���ҵ��������ұ�״̬���е�SQL��ɫͼ�꣬˫���㿪���ڷ�����ѡ��Sql Server Agent��Ȼ�������м�ͷ��ѡ���·��ĵ�����OSʱ�Զ���������<BR>
        8�����ʱ�����ݿ�ƻ��Ѿ��ɹ��������ˣ�������������������ý����Զ����� <BR>
        <BR>
        �޸ļƻ���<BR>
        1������ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����-->����-->���ݿ�ά���ƻ�-->�򿪺�ɿ������趨�ļƻ������Խ����޸Ļ���ɾ������ <BR>
        <BR>
        <B>�塢���ݵ�ת�ƣ��½����ݿ��ת�Ʒ�������</B><BR>
        <BR>
        һ������£����ʹ�ñ��ݺͻ�ԭ����������ת�����ݣ�����������£������õ��뵼���ķ�ʽ����ת�ƣ�������ܵľ��ǵ��뵼����ʽ�����뵼����ʽת������һ�����þ��ǿ������������ݿ���Ч�������������С�����������ݿ�Ĵ�С��������Ĭ��Ϊ����SQL�Ĳ�����һ�����˽⣬��������еĲ��ֲ��������⣬������ѯ���������Ա���߲�ѯ��������<BR>
        1����ԭ���ݿ�����б����洢���̵�����һ��SQL�ļ���������ʱ��ע����ѡ����ѡ���д�����ű��ͱ�д�����������Ĭ��ֵ�ͼ��Լ���ű�ѡ��<BR>
        2���½����ݿ⣬���½����ݿ�ִ�е�һ������������SQL�ļ�<BR>
        3����SQL�ĵ��뵼����ʽ���������ݿ⵼��ԭ���ݿ��е����б�����<BR>
            </blockquote></td>
          </tr>
        </table>
<%
end sub
%><!--#include file="Admin_copy.asp"-->
</body>
</html>