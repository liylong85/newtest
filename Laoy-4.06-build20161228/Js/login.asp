<!--#include file="../inc/conn.asp"-->
<%
If IsUser<>1 then
%>
document.writeln("		<form action=\"<%=SitePath%>User\/Userlogin.asp?action=login\" method=\"post\" name=loginForm>");
document.writeln("		<div class=\"loginForm\">�û�����<input name=\"Username\" class=\"borderall\" type=\"text\" size=\"15\" style=\"width:80px;height:15px;\">");
document.writeln("		���룺<input name=\"PassWord\" class=\"borderall\" type=\"password\" maxlength=\"16\" size=\"15\" style=\"width:80px;height:15px;\">");
document.writeln("		<input type=\"checkbox\" name=\"CookieDate\" id=\"CookieDate\" style=\"border:0;\"\/>����<\/div>");
document.writeln("		<div class=\"loginSelect\"><input id=\"loginBtn\" type=\"submit\" value=\"��¼\">");

//�������ҪQQ��¼���ܣ���������һ�е���ǰ�����˫б�ܣ�����//����
document.writeln("		<a id=\"loginQq\" href=\"<%=SitePath%>api\/qq\/redirect_to_login.asp\">��qq��¼<\/a>");

document.writeln("		<input id=\"loginBtn\" type=\"button\" value=\"ע��\" onClick=\"window.location.href=\'<%=SitePath%>User\/userreg.asp\'\"><\/div>");
document.writeln("		<\/form>");
<%else%>
document.writeln("<div style=\"margin-left:10px;line-height:20px;\">��ӭ�㣺<%=username%>��<%=moneyname%>��<b><%=mymoney%><\/b>���ȼ���<%=UserGroupInfo(LaoYdengji,0)%>��<a href=\"<%=SitePath%>User\/UserAdd.asp?action=useredit\">�޸�����<\/a>��<a href=\"<%=SitePath%>User\/UserAdd.asp?action=add\">��������<\/a>��<a href=\"<%=SitePath%>User\/UserLogin.asp?action=logout\">�˳�<\/a><\/div>");
<%end if%>