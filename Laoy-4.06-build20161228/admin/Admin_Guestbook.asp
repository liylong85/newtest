<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/Function_Page.asp"-->
<!--#include file="Admin_check.asp"-->
<%
Call chkAdmin(5)
%><html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��վ��̨����</title>
<LINK href="images/Admin_css.css" type=text/css rel=stylesheet>
<script src="js/admin.js"></script>
</head>

<body>
<table width="95%" border="0" cellspacing="2" cellpadding="3"  align=center class="admintable" style="margin-bottom:5px;">
    <tr><form name="form1" method="get" action="Admin_Guestbook.asp">
      <td height="25" bgcolor="f7f7f7">���ٲ��ң�
        <SELECT onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')"  size="1" name="s">
        <OPTION value="" selected>-=��ѡ��=-</OPTION>
        <OPTION value="?s=all">��������</OPTION>
        <OPTION value="?s=yn1">���������</OPTION>
        <OPTION value="?s=yn0">δ�������</OPTION>
      </SELECT>      </td>
      <td bgcolor="f7f7f7">
        <a href="?hits=1"></a>
        <input name="keyword" type="text" id="keyword" value="<%=request("keyword")%>" class="s26">
        <input name="Submit2" type="submit" class="bnt" value="����">
        </td>
      
    </form>
    </tr>
</table>
<%
	if request("action") = "add" then 
		call add()
	elseif request("action")="edit" then
		call edit()
	elseif request("action")="savenew" then
		call savenew()
	elseif request("action")="savedit" then
		call savedit()
	elseif request("action")="del" then
		call del()
	elseif request("action")="delAll" then
		call delAll()
	else
		call List()
	end if

sub List()
%>
<form name="myform" method="POST" action="Admin_Guestbook.asp?action=delAll">
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="6" align=left class="admintitle">�����б�</td>
</tr>
    <tr bgcolor="#f1f3f5" style="font-weight:bold;">
    <td width="5%" height="30" align="center" class="ButtonList">&nbsp;</td>
    <td width="47%" align="center" class="ButtonList">��������</td>
    <td width="17%" align="center" class="ButtonList">������</td>
    <td width="16%" height="25" align="center" class="ButtonList">����ʱ��</td>
    <td width="15%" height="25" align="center" class="ButtonList">����</td>    
    </tr>
<%
page=request("page")
s=Request("s")
id=request("id")
keyword=request("keyword")
Set mypage=new xdownpage
mypage.getconn=conn
mysql="select * from "&tbname&"_Guestbook"
	if s="yn0" then
	mysql=mysql&" Where yn=0"
	elseif s="yn1" then
	mysql=mysql&" Where yn=1"
	elseif keyword<>"" then
	mysql=mysql&" Where Content like '%"&keyword&"%'"
	End if
mysql=mysql&" order by id desc"
mypage.getsql=mysql
mypage.pagesize=15
set rs=mypage.getrs()
for i=1 to mypage.pagesize
    if not rs.eof then 
%>
    <tr>
    <td height="25" align="center" bgcolor="f7f7f7"><input type="checkbox" value="<%=rs("ID")%>" name="ID" onClick="unselectall(this.form)" style="border:0;"></td>
    <td height="25" bgcolor="f7f7f7"><%=left(LoseHtml(rs("Content")),30)%>...</td>
    <td height="25" align="center" bgcolor="f7f7f7"><%=rs("UserName")%></td>
    <td height="25" align="center" bgcolor="f7f7f7"><span class="td"><%=rs("AddTime")%></span></td>
    <td align="center" bgcolor="f7f7f7"><%if rs("yn")=0 then Response.Write("<font color=red>δ��</font>") else Response.Write("����") end if%>|<a href="?action=edit&id=<%=rs("ID")%>">�ظ�</a>|<a href="?action=del&id=<%=rs("ID")%>">ɾ��</a></td>
    </tr>
<%
        rs.movenext
    else
         exit for
    end if
next
%>
<tr><td align="center" bgcolor="f7f7f7"><input name="Action" type="hidden"  value="Del"><input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox" style="border:0"></td>
  <td colspan="5" bgcolor="f7f7f7"><input name="Del" type="submit" class="bnt" id="Del" value="ɾ��">
    <input name="Del" type="submit" class="bnt" id="Del" value="����δ��">

    <input name="Del" type="submit" class="bnt" id="Del" value="�������"></td>
  </tr><tr><td bgcolor="f7f7f7" colspan="6">
<div id="page">
	<ul style="text-align:left;">
    <%=mypage.showpage()%>
    </ul>
</div>
</td></tr></table>
</form>
<%
	rs.close
	set rs=nothing
end sub

sub del()
	id=LaoYRequest(request("id"))
	set rs=conn.execute("delete from "&tbname&"_Guestbook where id="&id)
	Response.write"<script>alert(""ɾ���ɹ���"");location.href=""Admin_Guestbook.asp"";</script>"
end sub

sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from "&tbname&"_Guestbook where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form name="myform" action="?action=savedit" method=post>
<tr> 
    <td colspan="5" class="admintitle">�޸�����</td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">������</td>
  <td colspan=4 bgcolor="#f7f7f7" class=td><input name="UserName" type="text" class="inputbg" id="UserName" value="<%=rs("UserName")%>" size="30"></td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">AddIP</td>
  <td colspan=4 bgcolor="#f7f7f7" class=td><%=rs("AddIP")%>������ʱ�䣺<%=rs("AddTime")%></td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">����</td>
  <td colspan=4 bgcolor="#f7f7f7" class=td><textarea name="Content" cols="80" rows="10" id="Content"><%=rs("Content")%></textarea></td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">�ظ�</td>
  <td colspan=4 bgcolor="#f7f7f7" class=td><textarea name="reContent" cols="80" rows="10" id="reContent"><%=rs("reContent")%></textarea></td>
</tr>
<tr> 
<td width="20%"></td>
<td colspan=4 class=td><input name="id" type="hidden" value="<%=rs("ID")%>"><input name="Submit" type="submit" class="bnt" value="�� ��">  </td>
</tr>
</form>
</table>
<%
rs.close
set rs=nothing
end sub

sub savedit()
	id=LaoYRequest(request.form("id"))
	UserName			=trim(request.form("UserName"))
	Content				=request.form("Content")
	reContent			=request.form("reContent")
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from "&tbname&"_Guestbook where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
	
		rs("UserName")			=UserName
		rs("Content")			=Content
		rs("reContent")			=reContent
		rs("ReTime")			=Now()
		rs.update
		Response.write"<script>alert(""��ϲ,�޸ĳɹ���"");location.href=""Admin_Guestbook.asp"";</script>"
	else
		Response.write"<script>alert(""�޸Ĵ���"");location.href=""Admin_Guestbook.asp"";</script>"
	end if
	rs.close
	set rs=nothing
end sub

Sub delAll
ID=Trim(Request("ID"))
If ID="" Then
	  Response.Write("<script language=javascript>alert('��ѡ��!');history.back(1);</script>")
	  Response.End
ElseIf Request("Del")="����δ��" Then
   set rs=conn.execute("Update "&tbname&"_Guestbook set yn = 0 where ID In(" & ID & ")")
   Response.Write("<script>alert(""�����ɹ���"");location.href=""Admin_Guestbook.asp"";</script>")
ElseIf Request("Del")="�������" Then
   set rs=conn.execute("Update "&tbname&"_Guestbook set yn = 1 where ID In(" & ID & ")")
   Response.Write("<script>alert(""�����ɹ���"");location.href=""Admin_Guestbook.asp"";</script>")
ElseIf Request("Del")="ɾ��" Then
	set rs=conn.execute("delete from "&tbname&"_Guestbook where ID In(" & ID & ")")
   	Response.write"<script>alert(""ɾ���ɹ���"");location.href=""Admin_Guestbook.asp"";</script>"
End If
End Sub
%>
<!--#include file="Admin_copy.asp"-->
</body>
</html>