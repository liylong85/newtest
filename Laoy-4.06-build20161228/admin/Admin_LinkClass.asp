<!--#include file="../Inc/conn.asp"-->
<!--#include file="admin_check.asp"-->
<%
Call chkAdmin(8)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="images/Admin_css.css" type=text/css rel=stylesheet>

</head>

<body>
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
	else
		call List()
	end if
 
sub List()
   Dim Sqlp,Rsp,TempStr
%>
<table border="0" cellspacing="2" cellpadding="3"  align="center" class="admintable">
<tr> 
  <td colspan="7" align=left class="admintitle">�����б���[<a href="?action=add">��������</a>]��[<a href="Admin_Link.asp?action=add">��������</a>]</td>
</tr>
  <tr align="center"> 
    <td width="36%" class="ButtonList">��������</td>
    <td width="10%" class="ButtonList">ID</td>
    <td width="18%" class="ButtonList">����</td>
    <td width="36%" class="ButtonList">�� ��</td>
  </tr>
<%
   Sqlp ="select * from "&tbname&"_LinkClass order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<tr><td colspan=""6"">û�з���</td></tr>")
   Else
   NoI=0
      Do while not Rsp.Eof   
	NoI=NoI+1
%>
    <tr bgcolor="#f1f3f5" onMouseOver="this.style.backgroundColor='#EAFCD5';this.style.color='red'" onMouseOut="this.style.backgroundColor='';this.style.color=''">
    <td height="25" class="tdleft"><%=rsp("LinkName")%></td>
    <td height="25" align="center" class="tdleft"><%=rsp("ID")%></td>
    <td height="25" align="center"><%=rsp("Num")%></td>
    <td width="36%" align="center"><a href="?action=edit&id=<%=rsp("ID")%>">�༭</a> | <a href="?action=del&id=<%=rsp("ID")%>" onClick="JavaScript:return confirm('ɾ������ͬʱ��ɾ���÷����µ��������ӣ�ȷ����')">ɾ��</a></td>
  </tr>
<%
      Rsp.Movenext   
      Loop   
   End if
   rsp.close
   set rsp=nothing
%>  
</table>
<%
end sub

sub add()
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="5" class="admintitle">���ӷ���</th></tr>
<form action="?action=savenew" method=post>
<tr>
<td width="20%" class=b1_1>��������</td>
<td colspan=4 class=b1_1><input type="text" name="LinkName" size="30"></td>
</tr>
<tr> 
<td width="20%" class=b1_1>�š�����</td>
<td colspan=4 class=b1_1><input name="num" type="text" value="10" size="4" maxlength="3"></td>
</tr>
<tr> 
<td width="20%" class=b1_1></td>
<td class=b1_1 colspan=4><input name="Submit" type="submit" class="bnt" value="�� ��"></td>
</tr></form>
</table>
<%
end sub

sub del()
	id=request("id")
		set rs=conn.execute("delete from "&tbname&"_LinkClass where id="&id)
		set rs=conn.execute("delete from "&tbname&"_Link where ClassID In(" & ID & ")")
	Response.write"<script>alert(""ɾ���ɹ���"");location.href=""Admin_LinkClass.asp"";</script>"
end sub

sub savenew()
	if trim(request.form("LinkName"))="" then
		Call Alert ("����д��������","-1")
	end if
	LinkName=trim(request.form("LinkName"))
	num=trim(request.form("num"))
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from "&tbname&"_LinkClass where LinkName='"& LinkName &"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.AddNew 
		rs("LinkName")		=LinkName
		rs("num")			=num	
		rs.update
		Call Alert ("��ϲ,���ӳɹ���","Admin_LinkClass.asp")
	else
		Call Alert ("����ʧ�ܣ��÷����Ѿ����ڣ�","Admin_LinkClass.asp")
	end if
	rs.close
	set rs=nothing
end sub

sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from "&tbname&"_LinkClass where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form action="?action=savedit" method=post>
<tr> 
    <td colspan="5" class="admintitle">�޸ķ���</td>
</tr>
<tr> 
<td width="20%" class="b1_1">��������</td>
<td colspan=4 class=b1_1><input type="text" name="LinkName" size="30" value="<%=rs("LinkName")%>"></td>
</tr>
<tr>
  <td class="b1_1">�š�����</td>
  <td colspan=4 class=b1_1><input name="Num" type="text" id="Num" value="<%=rs("Num")%>" size="4" maxlength="3"></td>
</tr>
<tr> 
<td width="20%" class="b1_1"></td>
<td class=b1_1 colspan=4><input name="id" type="hidden" value="<%=rs("ID")%>"><input name="Submit" type="submit" class="bnt" value="�� ��"></td>
</tr>
</form>
</table>
<%
rs.close
set rs=nothing
end sub

sub savedit()
	if trim(request.form("LinkName"))="" then
		Call Alert ("����д��������","-1")
	end if
	ID=trim(request.form("ID"))
	LinkName=trim(request.form("LinkName"))
	num=trim(request.form("num"))
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from "&tbname&"_LinkClass where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
		rs("LinkName")		=LinkName
		rs("num")			=num
		
		rs.update
		Call Alert ("��ϲ,�޸ĳɹ���","Admin_LinkClass.asp")
	else
		Call Alert ("�޸�ʧ�ܣ�","Admin_LinkClass.asp")
	end if
	rs.close
	set rs=nothing
end sub
%>
<!--#include file="Admin_copy.asp"-->
</body>
</html>