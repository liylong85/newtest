<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/md5.asp"-->
<%
If useroff=0 then Call Alert("��վĿǰ�Ѿ��رջ�Ա����","../")
	if request("action") = "login" then 
		call login()
	elseif request("action")="logout" then
		call logout()
	elseif request("action")="list" then
		call list()
	else
		call list()
	end if

Sub logout()
Response.Cookies("Yao").path=SitePath
Response.Cookies("Yao")("UserName")=""
Response.Cookies("Yao")("UserPass")=""
Response.Cookies("Yao")("ID")=""
Response.Cookies("Yao")("LaoYRndPassword")=""
Response.Redirect SitePath
End Sub

Sub login
UserName = CheckStr(trim(request.form("UserName")))
PassWord = md5(replace(trim(request.form("PassWord")),"'",""),16)
RndPassword=md5("l"&"a"&"o"&"y"&RndNumber(1,9999999999),32)
CookieDate=trim(request("CookieDate"))
if UserName="" then
	Call Alert ("�������û���","-1")
end if
if password="" then
	Call Alert ("����������","-1")
End if

Set rst = Server.CreateObject("ADODB.Recordset")
rst.Open "Select [username],[password] From ["&tbname&"_User] where password='" &password&"' and username='" &UserName&"'" , conn, 3,3
if rst.bof then
	Call Alert ("�û������������!","-1")
End if
rst.close
set rst=nothing

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select [ID],[username],[password],[LastIP],[LastTime],[yn],[RndPassword] From ["&tbname&"_User] where username='" &UserName&"'", conn, 3,3
if not rs.bof then
	If rs("yn")=0 then
		Call Alert ("�����ʺŻ�û��ͨ�����","-1")
	End if
	rs("LastIP")=GetIP
	rs("LastTime")=Now()
	rs("RndPassword")=RndPassword
	rs.update
	Response.Cookies("Yao").path=SitePath
		If CookieDate<>"" then
		Response.Cookies("Yao").Expires=Date+30
		End if
     	Response.Cookies("Yao")("UserName")=username
     	Response.Cookies("Yao")("UserPass")=PassWord
		Response.Cookies("Yao")("RndPassword")=RndPassword
      	Response.Cookies("Yao")("ID")=rs("ID")
end if
rs.close
set rs=nothing
Response.redirect Request.ServerVariables("HTTP_REFERER")
End Sub

Sub list
If LaoYID>0 then
Response.Redirect SitePath
End if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="keywords" content="<%=Sitekeywords%>" />
<link href="<%=SitePath%>images/css<%=Css%>.css" type=text/css rel=stylesheet>

<script type="text/javascript" src="<%=SitePath%>js/main.asp"></script>
<title>�û���¼</title>
</head>
<body>
<div class="mwall">
<%=Head%>
<%=Menu%><div class="mw">
<div style='margin:100px auto;'>
  <form action="<%=SitePath%>User/Userlogin.asp?action=login" method="post" name=loginForm>
<table width="300" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30" colspan="3" align="left">����û�е�½�����¼����<a href="userreg.asp"><u>������ע��</u></a></td>
    </tr>
  <tr>
    <td width="98" height="30">�û�����</td>
    <td colspan="2" align="left"><input name="Username" class="borderall" type="text" size="15" style="width:100px;height:15px;"></td>
  </tr>
  <tr>
    <td height="30">�ܡ��룺</td>
    <td colspan="2" align="left"><input name="PassWord" class="borderall" type="password" size="15" style="width:100px;height:15px;"></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="right"><input type="submit" name="Submit" value="��¼" class="borderall" style="height:19px;"/></td>
    <td width="179" align="left" style="padding-left:10px;"><a href="<%=sitepath%>api/qq/redirect_to_login.asp" id="loginQq">��QQ��¼</a></td>
  </tr>
</table>
</form>
</div>
</div>
<%=Copy%>
</div>
</body>
</html>
<%
End Sub
%>