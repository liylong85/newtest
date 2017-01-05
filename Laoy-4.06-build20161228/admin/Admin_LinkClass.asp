<!--#include file="../Inc/conn.asp"-->
<!--#include file="admin_check.asp"-->
<%
Call chkAdmin(8)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
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
  <td colspan="7" align=left class="admintitle">分类列表　[<a href="?action=add">分类添加</a>]　[<a href="Admin_Link.asp?action=add">链接添加</a>]</td>
</tr>
  <tr align="center"> 
    <td width="36%" class="ButtonList">分类名称</td>
    <td width="10%" class="ButtonList">ID</td>
    <td width="18%" class="ButtonList">排序</td>
    <td width="36%" class="ButtonList">操 作</td>
  </tr>
<%
   Sqlp ="select * from "&tbname&"_LinkClass order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<tr><td colspan=""6"">没有分类</td></tr>")
   Else
   NoI=0
      Do while not Rsp.Eof   
	NoI=NoI+1
%>
    <tr bgcolor="#f1f3f5" onMouseOver="this.style.backgroundColor='#EAFCD5';this.style.color='red'" onMouseOut="this.style.backgroundColor='';this.style.color=''">
    <td height="25" class="tdleft"><%=rsp("LinkName")%></td>
    <td height="25" align="center" class="tdleft"><%=rsp("ID")%></td>
    <td height="25" align="center"><%=rsp("Num")%></td>
    <td width="36%" align="center"><a href="?action=edit&id=<%=rsp("ID")%>">编辑</a> | <a href="?action=del&id=<%=rsp("ID")%>" onClick="JavaScript:return confirm('删除分类同时会删除该分类下的所有链接！确定？')">删除</a></td>
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
  <td colspan="5" class="admintitle">添加分类</th></tr>
<form action="?action=savenew" method=post>
<tr>
<td width="20%" class=b1_1>分类名称</td>
<td colspan=4 class=b1_1><input type="text" name="LinkName" size="30"></td>
</tr>
<tr> 
<td width="20%" class=b1_1>排　　序</td>
<td colspan=4 class=b1_1><input name="num" type="text" value="10" size="4" maxlength="3"></td>
</tr>
<tr> 
<td width="20%" class=b1_1></td>
<td class=b1_1 colspan=4><input name="Submit" type="submit" class="bnt" value="添 加"></td>
</tr></form>
</table>
<%
end sub

sub del()
	id=request("id")
		set rs=conn.execute("delete from "&tbname&"_LinkClass where id="&id)
		set rs=conn.execute("delete from "&tbname&"_Link where ClassID In(" & ID & ")")
	Response.write"<script>alert(""删除成功！"");location.href=""Admin_LinkClass.asp"";</script>"
end sub

sub savenew()
	if trim(request.form("LinkName"))="" then
		Call Alert ("请填写分类名称","-1")
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
		Call Alert ("恭喜,添加成功！","Admin_LinkClass.asp")
	else
		Call Alert ("添加失败，该分类已经存在！","Admin_LinkClass.asp")
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
    <td colspan="5" class="admintitle">修改分类</td>
</tr>
<tr> 
<td width="20%" class="b1_1">分类名称</td>
<td colspan=4 class=b1_1><input type="text" name="LinkName" size="30" value="<%=rs("LinkName")%>"></td>
</tr>
<tr>
  <td class="b1_1">排　　序</td>
  <td colspan=4 class=b1_1><input name="Num" type="text" id="Num" value="<%=rs("Num")%>" size="4" maxlength="3"></td>
</tr>
<tr> 
<td width="20%" class="b1_1"></td>
<td class=b1_1 colspan=4><input name="id" type="hidden" value="<%=rs("ID")%>"><input name="Submit" type="submit" class="bnt" value="提 交"></td>
</tr>
</form>
</table>
<%
rs.close
set rs=nothing
end sub

sub savedit()
	if trim(request.form("LinkName"))="" then
		Call Alert ("请填写分类名称","-1")
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
		Call Alert ("恭喜,修改成功！","Admin_LinkClass.asp")
	else
		Call Alert ("修改失败！","Admin_LinkClass.asp")
	end if
	rs.close
	set rs=nothing
end sub
%>
<!--#include file="Admin_copy.asp"-->
</body>
</html>