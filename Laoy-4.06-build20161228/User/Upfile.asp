<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/UpLoadClass.asp"-->
<%
Server.ScriptTimeOut=5000
If UserGroupInfo(LaoYdengji,5)<>1 or IsUser<>1 then
Call Alert("�ϴ������ѹر�!","-1")
End if
maxfile=150 '�ϴ��ļ���С����:��λkb,1024����1M
uploadpath=SitePath&SiteUp&"/" & Year(Now) & right("0" & Month(Now), 2) & "/"
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
	Response.Write("{""error"":1,""message"":""�벻Ҫ�Ƿ��ύ����!""}")
	Response.end
end if

dim request2,fpath
dim upload,file,formName,filename,fileExt,id,imgTitle,imgWidth,imgHeight,imgBorder,ming,txt
set request2=New UpLoadClass
CreateFolder(uploadpath&"index.html")
request2.SavePath=uploadpath
request2.AutoSave=2
request2.Open()
id=replace(trim(request2.form("id")),"'","")
imgTitle=replace(trim(request2.form("imgTitle")),"'","")
imgWidth=replace(trim(request2.form("imgWidth")),"'","")
imgHeight=replace(trim(request2.form("imgHeight")),"'","")
imgBorder=replace(trim(request2.form("imgBorder")),"'","")
fpath=trim(request2.form("imgFile"))

request2.MaxSize=1024*maxfile
if request2.Save("imgFile",0) then
	savefilename=request2.Form("imgFile")
	
	If IsAspJpeg=1 then
	Dim RV_img 
	RV_img=uploadpath&savefilename
	Call laoy_draw(RV_img)
	end if
	
	error1=0
else
	savefilename=""
	error1=1
	laoyerror=request2.Error
	Select case laoyerror
		case "1"
			message="ͼƬ���Ϊ"&maxfile&"K!"
		case "2"
			message="�ļ����Ͳ���ȷ"
		case else
			message="�ϴ����ɹ�!"
	end select
end if
set request2=nothing

if savefilename<>"" then%>
  {"error":<%= error1 %>,"url":"<%= uploadpath&savefilename %>"}
<% Else %>
  {"error":<%= error1 %>,"message":"����<%= message %>"}
<% End If %>