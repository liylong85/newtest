<!--#include file="Inc/conn.asp"-->
<!--#include file="Inc/Function_Page.asp"-->
<%If laoyvip then Response.Redirect sitepath&"bbs/"%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="keywords" content="<%=Sitekeywords%>" />
<meta name="description" content="<%=Sitedescription%>" />
<link href="<%=SitePath%>images/css<%=Css%>.css" type=text/css rel=stylesheet>
<script type="text/javascript" src="<%=SitePath%>js/main.asp"></script>
<title>���Ա� <%=SiteTitle%></title>
</head>
<body>
<div class="mwall">
<%=Head%>
<%=Menu%><div class="mw">
	<div class="dh">
		<%=search%>�����ڵ�λ�ã�<a href="Index.asp">��ҳ</a> >> <a href="guestbook.asp">�鿴����</a> <a href="?ac=add"><font style="color:#ff0000;">ǩд����</font></a>
    </div>
	<div id="nw_left">
		<div id="web2l">
        	<h6><%If request("ac")="add" then Response.Write("ǩд����") else Response.Write("�鿴����") end if%></h6>
			<div id="content">
				<%If request("ac")="" then%><ul>
<%
Set mypage=new xdownpage
mypage.getconn=conn
mypage.getsql="select * from "&tbname&"_Guestbook where yn=1 order by id desc"
mypage.pagesize=10
set rs=mypage.getrs()
NoI=0
for i=1 to mypage.pagesize
    if not rs.eof then
NoI=NoI+1
%>
<li>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="40%" height="25" valign="middle" class="bgf7f7f7"><font color=blue><%=NoI%>.</font><%If year(rs("AddTime"))&"-"&month(rs("AddTime"))&"-"&day(rs("AddTime"))=""&date&"" then Response.Write("<img src=""images/news.gif"" align=absMiddle title=""������...""> ") end if%><font color=red><%=rs("Title")%></font></td>
    <td class="bgf7f7f7"><%=left(rs("UserName"),15)%>&nbsp;<%=rs("AddTime")%> IP��<a href="http://www.laoy.net/Other/IP.asp?IP=<%=iparray(rs("AddIP"))%>" target="_blank"><%=iparray(rs("AddIP"))%></a></td>
  </tr>
  <tr>
    <td colspan="2" class="gcontent"><%=dvHTMLEncode(rs("Content"))%></td>
    </tr>
<%if rs("ReContent")<>"" then%>
  <tr>
    <td colspan="2" style="padding:5px 20px;line-height:20px;font-size:13px;color:#174BAF"><hr><font color=red>����Ա�ظ���</font><%=rs("ReContent")%><br><font color="#cccccc">(�ظ�ʱ�䣺<%=rs("ReTime")%>)</font></td>
    </tr>
<%end if%>
</table>
</li>
  <%
        rs.movenext
    else
         exit for
    end if
next
  %>
</ul>
<div id="clear"></div>
<div id="page">
	<ul>
    <%=mypage.showpage()%>
    </ul>
</div><%elseif request("ac")="add" then%>
	<script language=javascript>
function chk()
{
	if(document.form.title.value == "" || document.form.title.value.length > 15)
	{
	alert("�����ύ���ԣ�������Ա���Ϊ�ջ����15���ַ���");
	document.form.title.focus();
	document.form.title.select();
	return false;
	}
	if(document.form.content.value == "")
	{
	alert("����д�������ݣ�");
	document.form.content.focus();
	document.form.content.select();
	return false;
	}
	if(document.form.code.value == "")
	{
	alert("����д��֤�룡");
	document.form.code.focus();
	document.form.code.select();
	return false;
	}
return true;
}
</script>
<form onSubmit="return chk();" method="post" name="form" action="?ac=post">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#ffffff">
    <tr>
      <td height="30" class="adgs">������
      <input name="UserName" type="text" id="UserName" maxlength="10" value="<%If IsUser=1 then Response.Write(""&UserName&"") else Response.Write(""&iparray(GetIP)&"") End if%>" Readonly style="width:150px;border:1px solid #ccc;"></td>
    </tr>
    <tr>
      <td height="30" class="adgs">���⣺
      <input name="title" type="text" id="title" maxlength="15" style="width:400px;border:1px solid #ccc;"></td>
    </tr>
    <tr>
      <td height="15" class="adgs">���ݣ�
      <textarea name="content" cols="30" rows="6" id="content" style="width:400px;border:1px solid #ccc;margin:0;padding:0;height:100px;font-size:12px;line-height:120%;"></textarea></td>
    </tr>
	<tr>
      <td height="30" class="adgs">��֤�룺
      <input name="code" type="text" id="code" size="8" maxlength="5" style="border:1px solid #ccc;"/>
        <img src="Inc/code.asp" border="0" alt="�����������ˢ����֤��" style="cursor : pointer;" onclick="this.src='../Inc/code.asp?y='+Math.random()"/>
	</td>
    </tr>
    <tr>
      <td height="30" align="center"><input type="submit" name="Submit" value=" �� �� "></td>
    </tr>
  </table>
</form>
	<%elseif request("ac")="post" then
	dim UserName,Title,Content
	UserName = 	CheckStr(trim(request.form("UserName")))
	Title = 	CheckStr(trim(request.form("Title")))
	Content = 	request.form("Content")
	mycode = trim(request.form("code"))
	if mycode<>Session("getcode") then
		Call Alert("��������ȷ����֤�룡",-1)
	end if
  	If ChkSB(Content)=false Then
	Call Alert("�벻Ҫ�ҷ���Ϣ!",-1)
	end if
	If session("postgstime")<>"" then
		posttime8=DateDiff("s",session("postgstime"),Now())
  		if posttime8<yaopostgetime then
		posttime9=yaopostgetime-posttime8
		Call Alert("�벻Ҫ��������!",-1)
  		end if
	End if
	If Not Checkpost(True) Then Call Alert("��ֹ�ⲿ�ύ!","-1")
	set rsg = server.CreateObject ("adodb.recordset")
	sql="select * from "&tbname&"_Guestbook"
	rsg.open sql,conn,1,3
		if UserName="" or Title="" or Content="" then
			response.Redirect "Guestbook.asp?ac=add"
		end if

		rsg.AddNew 
		rsg("UserName")			=UserName
		rsg("Title")				=Title
		rsg("Content")			=Left(Content,200)
		rsg("AddIP")				=GetIP
		If bookoff=1 then
		rsg("yn")				=1
		else
		rsg("yn")				=0
		end if
		rsg.update
		Session("postgstime")=Now()
		If bookoff=0 then
			Call Alert ("��ϲ��,���Գɹ�,����Ҫ����Ա��˺������ʾ����!","index.asp")
		else
			Call Alert ("��ϲ��,���Գɹ�!","Guestbook.asp")
		end if
		rsg.close
		Set rsg = nothing
		end if%>
			</div>
		</div>
	</div>
	<div id="nw_right">
		<%Echo ShowAD(3)%>
		<div id="web2r">
			<h5>��������</h5>
			<ul id="list10">
            	<%Call ShowArticle(0,10,5,"��",100,"no","Hits desc,ID desc",0,0,0)%>
            </ul>
  		</div>
        <div id="web2r">
			<h5>ͼƬ�Ƽ�</h5>
			<ul class="topimg">
                <%Call ShowImgArticle(ID,4,100,"no","DateAndTime desc,ID desc",60,60,1)%>
            </ul>
  		</div>
	</div>
</div>
<%=Copy%></div>
</body>
</html>