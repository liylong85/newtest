<!--#include file="../Inc/Conn.asp"-->
<!--#include file="../Inc/UpLoadClass.asp"-->
<%
If IsUser<>1 then
Call Alert ("û��Ȩ��","-1")
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Upload</title>

<style>
body {padding:0;margin:0;background:#fff;}
body,td {font-size:12px;}
</style>
<body>
<%
select case request("action")
case "simg":
call simg()
end select

sub simg()
if Request.QueryString("submit")="simg" then
uploadpath=SitePath&SiteUp&"/UserFace/"
uploadsize="1024"
uploadtype="jpg/gif/png/bmp"
Set Uprequest=new UpLoadClass
	CreateFolder sitepath&uploadpath&"/index.html"
    Uprequest.SavePath=uploadpath
    Uprequest.MaxSize=uploadsize*500
    Uprequest.FileType=uploadtype
    AutoSave=true
    Uprequest.open
  if Uprequest.form("file_Err")<>0  then
  select case Uprequest.form("file_Err")
  case 1:str="<div style=""padding-top:5px;padding-bottom:5px;""> <font color=blue>"&Uprequest.MaxSize/1024&"K [<a href='javascript:history.go(-1)'>�����ϴ�</a>]</font></div>"
  case 2:str="<div style=""padding-top:5px;padding-bottom:5px;""> <font color=blue>�ļ���ʽ���� [<a href='javascript:history.go(-1)']>�����ϴ�</a>]</font></div>"
  case 3:str="<div style=""padding-top:5px;padding-bottom:5px;""> <font color=blue>�ļ�̫���Ҹ�ʽ���� [<a href='javascript:history.go(-1)'>�����ϴ�</a>]</font></div>"
  end select
  response.write str
  else
	sql="Update "&tbname&"_User set UserFace = '"&CheckStr(Uprequest.Form("file"))&"' where ID= "&LaoYID
	conn.execute(sql)
  response.write "<div style=""margin-top:6px;""><font color=red>�ϴ��ɹ�</font>,ˢ�¸�ҳ�漴�ɿ����µ�ͷ��<a href='javascript:history.go(-1)'>�����ϴ�</a></div>"
  
  conn.close
  set conn=nothing
  end if

'Set Uprequest=nothing

'����֧�����ʱ������ˮӡ����ͼ
'If IsAspJpeg=1 then

Dim Jpeg,RV_img 
RV_img=Uprequest.SavePath&Uprequest.Form("file")

'����ͷ��
If right(RV_img,4)<>".gif" then
Dim S_Width,S_Height,H_Temp,W_Temp
S_Width=100
S_Height=100
Set Jpeg = Server.CreateObject("Persits.Jpeg") '����ʵ��
Path = Server.MapPath(RV_img) '����ͼƬ·��
Jpeg.Open Path '��ͼƬ

  If Jpeg.OriginalWidth>S_Width or Jpeg.OriginalHeight>S_Height Then
    H_Temp=S_Width*Jpeg.OriginalHeight/Jpeg.OriginalWidth '����[��]��ΪСͼ���ֵʱ,ȡ�õȱ����ߵĳߴ�.
    W_Temp=Jpeg.OriginalWidth*S_Height/Jpeg.OriginalHeight '����[��]��ΪСͼ���ֵʱ,ȡ�õȱ�����ĳߴ�.
    If W_Temp>S_Width Then           '�������ʱֵ��������ʱ: ��ȡ��Сͼ������ֵ,�߰�������ֵ����ó�
     Jpeg.Width=W_Temp
     Jpeg.Height=S_Height  
    Else                    '���ߵ���ʱֵ��������ʱ: ��ȡ��Сͼ�ߵ����ֵ,���ߵ����ֵ����ó�
	 Jpeg.Width =S_Width
     Jpeg.Height=H_Temp
    End If
	Jpeg.crop 0,0,S_Width,S_Height
  Else 
   Jpeg.Width=Jpeg.OriginalWidth
   Jpeg.Height=Jpeg.OriginalHeight
  End If

  Jpeg.Save Server.MapPath(Uprequest.SavePath&Uprequest.Form("file")) '����ͼƬ������
Jpeg.Close:Set Jpeg = Nothing
'end if
end if
end if

response.write "<form name=form action=?action=simg&submit=simg method=post enctype=multipart/form-data>"
response.write "<div style='text-align:left;'><input type=file name=file size=15  style=""width:250px;"" onfocus=""javascript:this.className='fbform1';"" onblur=""javascript:this.className='fbform';"">&nbsp;"
response.write "<input type=submit name=submit value=�ϴ�></div>"
response.write "</form>"
end sub
%>
</body>
</html>