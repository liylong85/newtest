<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/Function_Page.asp"-->
<!--#include file="Admin_check.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��վ��̨����</title>
<link href="images/Admin_css.css" type=text/css rel=stylesheet>

<script src="js/admin.js"></script>
</head>

<body>
<%
Call chkAdmin(4)
	if request("action") = "add" then 
		call add()
	elseif request("action")="edit" then
		call edit()
	elseif request("action")="savenew" then
		call savenew()
	elseif request("action")="savedit" then
		call savedit()
	elseif request("action")="yn1" then
		call yn1()
	elseif request("action")="yn2" then
		call yn2()
	elseif request("action")="del" then
		call del()
	elseif request("action")="delAll" then
		call delAll()
	else
		call List()
	end if

sub List()
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="4" align=left class="admintitle">��ǩ�б���[<a href="?action=add">����</a>]</td></tr>
    <tr bgcolor="#f1f3f5" style="font-weight:bold;">
    <td height="30" align="center" class="ButtonList">��ǩ����</td>
    <td width="23%" height="25" align="center" class="ButtonList">����ʱ��</td>
    <td height="25" align="center" class="ButtonList">��ǩID</td>
    <td height="25" align="center" class="ButtonList">����</td>    </tr>
<%
page=request("page")
Set mypage=new xdownpage
mypage.getconn=conn
mysql="select * from "&tbname&"_Label order by id desc "
mypage.getsql=mysql
mypage.pagesize=20
set rs=mypage.getrs()
for i=1 to mypage.pagesize
    if not rs.eof then
%>
    <tr bgcolor="#f1f3f5" onMouseOver="this.style.backgroundColor='#EAFCD5';this.style.color='red'" onMouseOut="this.style.backgroundColor='';this.style.color=''">
    <td height="25"><a href="?action=edit&id=<%=rs("ID")%>"><%=rs("Title")%></a></td>
    <td height="25" align="center"><%=rs("DateAndTime")%></td>
    <td width="7%" height="25" align="center"><%=rs("ID")%></td>
    <td width="24%" align="center"><a href="?action=edit&id=<%=rs("ID")%>">�༭</a></td>    </tr>
<%
        rs.movenext
    else
         exit for
    end if
next
%><tr><td bgcolor="f7f7f7" colspan="4" align="left">���ã�����Ҫ���õĵط����� &lt;%Echo Label(��ǩID)%&gt; ���ɡ�
<div id="page">
    <ul style="text-align:left;">
      <%=mypage.showpage()%>
    </ul>
  </div>
</td></tr></table>
<%
	rs.close
	set rs=nothing
end sub

sub add()
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form action="?action=savenew" name="myform" method=post>
<tr> 
    <td colspan="2" align=left class="admintitle">���ӱ�ǩ</td></tr>
<tr> 
<td width="20%" class="b1_1">����</td>
<td class="b1_1"><input name="Title" type="text" id="Title" size="40" maxlength="50"></td>
</tr>
<tr>
  <td valign="top" class="b1_1">����</td>
  <td class="b1_1"><textarea name="Content" cols="80" rows="10" id="Content"></textarea></td>
</tr>
<tr> 
<td width="20%" class="b1_1"></td>
<td class="b1_1"><input name="Submit" type="submit" class="bnt" value="�� ��"></td>
</tr>
</form>
</table>
<%
end sub

sub savenew()
	Title			=trim(request.form("Title"))
	Content			=request.form("Content")
	
	if Title="" or Content="" then
		Call Alert ("����д����","-1")
	end if
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from "&tbname&"_Label where Title='"&Title&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.AddNew 
		rs("Title")				=Title
		rs("Content")			=Content
		rs.update
		Response.write"<script>alert(""���ӳɹ���"");location.href=""Admin_Label.asp"";</script>"
	else
		Response.Write("<script language=javascript>alert('�ñ�ǩ�Ѵ���!');history.back(1);</script>")
	end if
	rs.close
	set rs=nothing
end sub

sub del()
	id=request("id")
	set rs=conn.execute("delete from "&tbname&"_Label where id="&id)
	Response.write"<script>alert(""ɾ���ɹ���"");location.href=""Admin_Label.asp"";</script>"
end sub

sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from "&tbname&"_Label where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form name="myform" action="?action=savedit" method=post>
<tr> 
    <td colspan="5" class="admintitle">�޸ı�ǩ</td></tr>
<tr>
  <td width="20%" bgcolor="#FFFFFF" class="b1_1">����</td>
  <td colspan=4 class=b1_1><input name="Title" type="text" value="<%=rs("Title")%>" size="40" maxlength="50"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">����</td>
  <td colspan=4 class=b1_1><textarea name="Content" cols="80" rows="10" id="Content"><%=rs("Content")%></textarea></td>
</tr>
<tr> 
<td width="20%" class="b1_1"></td>
<td colspan=4 class=b1_1><input name="id" type="hidden" value="<%=rs("ID")%>"><input name="Submit" type="submit" class="bnt" value="�� ��"></td>
</tr>
</form>
</table>
<%
rs.close
set rs=nothing
end sub

sub savedit()
	Dim Title
	id=request.form("id")
	Title			=trim(request.form("Title"))
	Content			=request.form("Content")
	
	if Title="" then
		Call Alert ("����д����","-1")
	end if
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from "&tbname&"_Label where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
	
		rs("Title")				=Title
		rs("Content")			=Content
		rs("DateAndTime")		=Now
		
		rs.update
		Response.write"<script>alert(""�޸ĳɹ���"");location.href=""Admin_Label.asp"";</script>"
	else
		Response.write"<script>alert(""�޸Ĵ���"");location.href=""Admin_Label.asp"";</script>"
	end if
	rs.close
	set rs=nothing
end sub
%>
<!--#include file="Admin_copy.asp"-->
</body>
</html>