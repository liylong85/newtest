<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/Function_Page.asp"-->
<!--#include file="Admin_check.asp"-->
<%
Call chkAdmin(7)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��վ��̨����</title>
<link href="images/Admin_css.css" type=text/css rel=stylesheet>

<script src="js/admin.js"></script>
</head>

<body>
<table width="95%" border="0" cellspacing="2" cellpadding="3"  align=center class="admintable" style="margin-bottom:5px;">
    <tr><form name="form1" method="get" action="Admin_Link.asp">
      <td height="25" bgcolor="f7f7f7">���ٲ��ң�
        <SELECT onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')"  size="1" name="s">
        <OPTION value="" selected>-=��ѡ��=-</OPTION>
        <OPTION value="?s=all">��������</OPTION>
        <OPTION value="?s=yn1">���������</OPTION>
        <OPTION value="?s=yn0">δ�������</OPTION>
        <OPTION value="?s=3">Logo����</OPTION>
        <OPTION value="?s=1">�Ѿ����ڵ�����</OPTION>
        <OPTION value="?s=2">������ڵ�����</OPTION>
        <OPTION value="?s=4">������ڵ�����</OPTION>
      </SELECT>      </td>
      <td align="center" bgcolor="f7f7f7">
        <a href="?hits=1"></a>
        <input name="keyword" type="text" id="keyword" value="<%=request("keyword")%>">
        <input type="submit" name="Submit2" value="����">
      
    </form></td>
      <td align="right" bgcolor="f7f7f7">��ת����
        <select name="ClassID" id="ClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
    <%
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from "&tbname&"_LinkClass order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">��ѡ�����</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">�������ӷ���</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """?ClassID=" & Rsp("ID") & """" & "")
		 If int(request("ClassID"))=Rsp("ID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">|-" & Rsp("LinkName") & "")
         Response.Write("</option>" ) 
      Rsp.Movenext   
      Loop   
   End if
   Rsp.close:Set Rsp=nothing
	%>
        </select></td>
    </tr>
</table>
<script language="javascript" type="text/javascript" src="<%=SitePath%>js/setdate/WdatePicker.js"></script>
<%
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
<form name="myform" method="POST" action="Admin_Link.asp?action=delAll">
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="8" align=left class="admintitle">�����б���[<a href="?action=add">����</a>]��[<a href="Admin_LinkClass.asp">���ӷ���</a>]</td>
</tr>
<tr bgcolor="#f1f3f5" style="font-weight:bold;">
	<td width="3%" height="30" align="center" class="ButtonList">&nbsp;</td>
    <td width="37%" height="30" align="center" class="ButtonList">��������</td>
    <td width="5%" align="center" class="ButtonList">����</td>
    <td width="12%" height="25" align="center" bgcolor="#f1f3f5" class="ButtonList">����ʱ��</td>
    <td width="12%" height="25" align="center" bgcolor="#f1f3f5" class="ButtonList">����ʱ��</td>
    <td width="5%" height="25" align="center" bgcolor="#f1f3f5" class="ButtonList">����</td>
    <td width="14%" align="center" bgcolor="#f1f3f5" class="ButtonList">��ϵQQ</td>
    <td width="12%" height="25" align="center" bgcolor="#f1f3f5" class="ButtonList">����</td>    
  </tr>
<%
If laoyvip then dlink=3 else dlink=2
page=request("page")
Hits=request("hits")
s=Request("s")
Articleclass=request("ClassID")
keyword=request("keyword")
NOI=0
Set mypage=new xdownpage
mypage.getconn=conn
mysql="select * from "&tbname&"_Link Where yn<>"&dlink&""
	if s="yn0" then
	mysql=mysql&" And yn=0"
	elseif s="yn1" then
	mysql=mysql&" And yn=1"
	elseif s="1" then
	If IsSqlDataBase = 1 Then
	mysql=mysql&" And datediff(dd,GetDate(),LastTime) <= 0"
	else
	mysql=mysql&" And datediff('d',Now(),LastTime) <= 0"
	end if
	elseif s="2" then
	If IsSqlDataBase = 1 Then
	mysql=mysql&" And datediff(dd,GetDate(),LastTime) = 1"
	else
	mysql=mysql&" And datediff('d',Now(),LastTime) = 1"
	end if
	elseif s="4" then
	If IsSqlDataBase = 1 Then
	mysql=mysql&" And datediff(dd,GetDate(),LastTime) = 2"
	else
	mysql=mysql&" And datediff('d',Now(),LastTime) = 2"
	end if
	elseif s="3" then
	mysql=mysql&" And LogoUrl<>''"
	elseif s="ishot" then
	mysql=mysql&" And ishot=1"
	elseif s="img" then
	mysql=mysql&" And images<>''"
	elseif s="url" then
	mysql=mysql&" And LinkUrl<>''"
	elseif s="user" then
	mysql=mysql&" And UserName<>''"
	elseif Articleclass<>"" then
	mysql=mysql&" And ClassID="&Articleclass&""
	elseif keyword<>"" then
	mysql=mysql&" And Title like '%"&keyword&"%' or LinkUrl like '%"&keyword&"%' or QQ like '%"&keyword&"%'"
	End if
mysql=mysql&" order by Num asc,id asc"
mypage.getsql=mysql
mypage.pagesize=100
set rs=mypage.getrs()
for i=1 to mypage.pagesize
    if not rs.eof then
NOI=NOI+1
%>
    <tr bgcolor="#f1f3f5" onMouseOver="this.style.backgroundColor='#EAFCD5';this.style.color='red'" onMouseOut="this.style.backgroundColor='';this.style.color=''">
    <td height="25" align="CENTER"><input type="checkbox" value="<%=rs("ID")%>" name="ID" onClick="unselectall(this.form)" style="border:0;">
      <input name="LinkID" type="hidden" id="LinkID" value="<%=rs("ID")%>"></td>
    <td height="25"><%=NOI%>.[<a href="?ClassID=<%=rs("ClassID")%>"><%=linkclassname(rs("ClassID"))%></a>]
      <input name="Title" type="text" id="Title" value="<%=rs("Title")%>" size="15">
      <input name="LinkUrl" type="text" id="LinkUrl" value="<%=rs("LinkUrl")%>" size="20">
      <%If rs("Logourl")<>"" then Response.Write("<font color=red>[ͼ]</font>") End if%></td>
    <td height="25" align="center"><input name="Moneys" type="text" id="Moneys" value="<%=rs("Moneys")%>" size="4" maxlength="4"></td>
    <td height="25" align="center"><input name="AddTimes" type="text" id="AddTimes" value="<%=rs("AddTime")%>" size="9"></td>
    <td height="25" align="center"><input name="LastTime" type="text" id="LastTime" value="<%=rs("LastTime")%>" size="9"><%If datediff("d",Now(),rs("LastTime"))<=0 then Response.Write("<font color=red>[�ѹ���]</font>") else:If datediff("d",Now(),rs("LastTime"))<=1 then Response.Write("<font color=red>[�������]</font>") End if%></td>
    <td height="25" align="center"><input name="Num" type="text" id="Num" value="<%=rs("Num")%>" size="4" maxlength="4"></td>
    <td height="25" align="center"><%if rs("qq")<>"" then%><a href="tencent://message/?uin=<%=rs("qq")%>&amp;Site=www.laoy.net&amp;Menu=yes"><img border=0 src=http://wpa.qq.com/pa?p=4:<%=rs("qq")%>:4 alt="���������ҷ���Ϣ" height=16 align="absmiddle" /></a><%end if%><input name="qq" type="text" id="qq" value="<%=rs("qq")%>" size="10"></td>
    <td align="center"><%If rs("yn")=1 then Response.Write("����") else Response.Write("<font color=red>δ��</font>") end if%>|<a href="?action=edit&id=<%=rs("ID")%>">�༭</a></td>    
  </tr>
<%
        rs.movenext
    else
         exit for
    end if
next
%>
	<tr bgcolor="#f1f3f5" onMouseOver="this.style.backgroundColor='#EAFCD5';this.style.color='red'" onMouseOut="this.style.backgroundColor='';this.style.color=''">
    <td height="25" align="center" bgcolor="#f1f3f5"><input name="Action" type="hidden"  value="Del">
      <input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox" style="border:0"></td>
    <td height="25" colspan="7" bgcolor="#f1f3f5"><font color=red>�ƶ�����</font>
      <select id="ytype" name="ytype">
        <%call Admin_ShowClass_Option()%>
      </select>
&nbsp;
<input name="Del" type="submit" class="bnt" id="Del" value="�ƶ�">
<input name="Del" type="submit" class="bnt" id="Del"  onClick="JavaScript:return confirm('ɾ����')" value="ɾ��">
<input name="Del" type="submit" class="bnt" id="Del" value="����δ��">
<input name="Del" type="submit" class="bnt" id="Del" value="�������">
<input name="Del" type="submit" class="bnt" id="Del" value="�༭">
������(<%=month(Now())%>��)���룺 <b><%=Mydb("Select Sum([Moneys]) From ["&tbname&"_Link] Where month(AddTime)=month("&SqlNowString&")",1)(0)%></b> Ԫ</td>
    </tr>
  <tr><td bgcolor="f7f7f7" colspan="8" align="left"><div id="page">
	<ul style="text-align:left;">
    <%=mypage.showpage()%>
    </ul>
</div>
</td></tr>
</table>
</form>
<%
rs.close
set rs=nothing
end sub

sub add()
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form action="?action=savenew" name="myform" method=post>
<tr> 
    <td colspan="3" align=left class="admintitle">�������� [<a href="Admin_link.asp">�����б�</a>]</td>
</tr>
<tr> 
<td width="20%" class="b1_1">����</td>
<td colspan="2" class="b1_1"><input name="Title" type="text" id="Title" size="40" maxlength="50"></td>
</tr>
<tr>
  <td class="b1_1">��������</td>
  <td colspan="2" class="b1_1"><select ID="ClassID" name="ClassID">
    <%call Admin_ShowClass_Option()%>
  </select></td>
</tr>
<tr>
  <td class="b1_1">���ӵ�ַ</td>
  <td colspan="2" class="b1_1"><input name="LinkUrl" type="text" id="LinkUrl" value="http://" size="40" maxlength="50"></td>
</tr>
<tr>
  <td class="b1_1">Logo��ַ</td>
  <td width="21%" class="b1_1"><input name="LogoUrl" type="text" id="LogoUrl" size="40" maxlength="250"></td>
  <td width="59%" class="b1_1"><iframe src="upload.asp?action=link" width="400" height="25" frameborder="0" scrolling="no"></iframe></td>
</tr>
<tr>
  <td class="b1_1">��ʼʱ��</td>
  <td colspan="2" class="b1_1"><input name='AddTime' type='text' class="borderall" id="AddTime" style="width:140px;" onFocus="WdatePicker({isShowClear:false,readOnly:true,minDate:'%y-%M-#{%d}',skin:'whyGreen'})" value="<%=date%>"/></td>
  </tr>
<tr>
  <td class="b1_1">����ʱ��</td>
  <td colspan="2" class="b1_1"><input name='LastTime' type='text' class="borderall" id="LastTime" style="width:140px;" onFocus="WdatePicker({isShowClear:false,readOnly:true,minDate:'%y-%M-#{%d}',skin:'whyGreen'})" value="<%=date+30%>"/></td>
</tr>
<tr>
  <td class="b1_1">����</td>
  <td colspan="2" class="b1_1"><input name="Num" type="text" id="Num" value="<%=iif(Mydb("Select Max([Num]) From ["&tbname&"_Link]",1)(0)+1>0,Mydb("Select Max([Num]) From ["&tbname&"_Link]",1)(0)+1,1)%>" size="5" maxlength="3"> <span class="note">ֵԽСԽ��ǰ</span></td>
</tr>
<tr>
  <td class="b1_1">��ϵQQ</td>
  <td colspan="2" class="b1_1"><input name="QQ" type="text" id="QQ" size="15" maxlength="10"></td>
</tr>
<tr>
  <td class="b1_1">����</td>
  <td colspan="2" class="b1_1"><input name="Moneys" type="text" id="Moneys" value="0" size="8" maxlength="5"></td>
</tr>
<tr>
  <td valign="top" class="b1_1">��ע</td>
  <td colspan="2" class="b1_1"><textarea name="Content" cols="50" rows="5" id="Content"></textarea></td>
</tr>
<tr> 
<td width="20%" class="b1_1"></td>
<td colspan="2" class="b1_1"><input name="Submit" type="submit" class="bnt" value="�� ��"></td>
</tr>
</form>
</table>
<%
end sub

sub savenew()
	Title			=trim(request.form("Title"))
	ClassID			=trim(request.form("ClassID"))
	LinkUrl			=trim(request.form("LinkUrl"))
	LogoUrl			=trim(request.form("LogoUrl"))
	AddTime			=trim(request.form("AddTime"))
	LastTime		=trim(request.form("LastTime"))
	Num				=trim(request.form("Num"))
	QQ				=trim(request.form("qq"))
	Content			=request.form("Content")
	Moneys			=request.form("Moneys")
	
	if Title="" or LinkUrl="" or ClassID="" or LastTime="" then
		Call Alert ("����д����","-1")
	end if
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from "&tbname&"_Link"
	rs.open sql,conn,1,3
		rs.AddNew 
		rs("Title")				=Title
		rs("ClassID")			=ClassID
		rs("LinkUrl")			=LinkUrl	
		rs("LogoUrl")			=LogoUrl
		rs("AddTime")			=AddTime
		rs("LastTime")			=LastTime
		rs("Num")				=Num
		rs("qq")				=qq
		rs("Content")			=Content
		rs("yn")				=1
		rs("Moneys")				=Moneys
		rs.update
		session("YaoLinkClassID")=ClassID
		Call Alert ("���ӳɹ�","Admin_Link.asp?action=add")
	rs.close
	set rs=nothing
end sub

sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from "&tbname&"_Link where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form action="?action=savedit" name="myform" method=post>
<tr> 
    <td colspan="3" align=left class="admintitle">�޸�����</td>
</tr>
<tr> 
<td width="20%" class="b1_1">����</td>
<td colspan="2" class="b1_1"><input name="Title" type="text" id="Title" value="<%=rs("Title")%>" size="40" maxlength="50"></td>
</tr>
<tr>
  <td class="b1_1">��������</td>
  <td colspan="2" class="b1_1"><select ID="ClassID" name="ClassID">
    <%
   Set Rsp=server.CreateObject("adodb.recordset") 
   Sqlp ="select * from "&tbname&"_LinkClass order by num"   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">��ѡ�����</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">�������ӷ���</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """" & Rsp("ID") & """" & "")
         If rs("ClassID")=Rsp("ID") Then
            Response.Write(" selected")
         End If
         Response.Write(">|-" & Rsp("LinkName") & "")			
         Response.Write("</option>" ) 
      Rsp.Movenext   
      Loop   
   End if
   %>
  </select></td>
</tr>
<tr>
  <td class="b1_1">���ӵ�ַ</td>
  <td colspan="2" class="b1_1"><input name="LinkUrl" type="text" id="LinkUrl" value="<%=rs("LinkUrl")%>" size="40" maxlength="50"></td>
</tr>
<tr>
  <td class="b1_1">Logo��ַ</td>
  <td width="21%" class="b1_1"><input name="LogoUrl" type="text" id="LogoUrl" value="<%=rs("LogoUrl")%>" size="40" maxlength="50"></td>
  <td width="59%" class="b1_1"><iframe src="upload.asp?action=link" width="400" height="25" frameborder="0" scrolling="no"></iframe></td>
</tr>
<tr>
  <td class="b1_1">��ʼʱ��</td>
  <td colspan="2" class="b1_1"><input name='AddTime' type='text' class="borderall" id="AddTime" style="width:140px;" onFocus="WdatePicker({isShowClear:false,readOnly:true,skin:'whyGreen'})" value="<%=rs("AddTime")%>"/></td>
</tr>
<tr>
  <td class="b1_1">����ʱ��</td>
  <td colspan="2" class="b1_1"><input name='LastTime' type='text' class="borderall" id="LastTime" style="width:140px;" onFocus="WdatePicker({isShowClear:false,readOnly:true,minDate:'%y-%M-#{%d}',skin:'whyGreen'})" value="<%=rs("LastTime")%>"/></td>
</tr>
<tr>
  <td class="b1_1">����</td>
  <td colspan="2" class="b1_1"><input name="Num" type="text" id="Num" value="<%=rs("Num")%>" size="5" maxlength="3"> 
  <span class="note">ֵԽСԽ��ǰ</span></td>
</tr>
<tr>
  <td class="b1_1">��ϵQQ</td>
  <td colspan="2" class="b1_1"><input name="QQ" type="text" id="QQ" value="<%=rs("qq")%>" size="15" maxlength="10"></td>
</tr>
<tr>
  <td class="b1_1">����</td>
  <td colspan="2" class="b1_1"><input name="Moneys" type="text" id="Moneys" value="<%=rs("Moneys")%>" size="8" maxlength="5"></td>
</tr>
<tr>
  <td valign="top" class="b1_1">��ע</td>
  <td colspan="2" class="b1_1"><textarea name="Content" cols="50" rows="5" id="Content"><%=rs("Content")%></textarea></td>
</tr>
<tr> 
<td width="20%" class="b1_1"><input name="ID" type="hidden" id="ID" value="<%=rs("ID")%>"></td>
<td colspan="2" class="b1_1"><input name="Submit" type="submit" class="bnt" value="�� ��"></td>
</tr>
</form>
</table>
<%
rs.close
set rs=nothing
end sub

sub savedit()
	ID				=trim(request.form("ID"))
	Title			=trim(request.form("Title"))
	ClassID			=trim(request.form("ClassID"))
	LinkUrl			=trim(request.form("LinkUrl"))
	LogoUrl			=trim(request.form("LogoUrl"))
	AddTime			=trim(request.form("AddTime"))
	LastTime		=trim(request.form("LastTime"))
	Num				=trim(request.form("Num"))
	QQ				=trim(request.form("qq"))
	Content			=request.form("Content")
	Moneys			=request.form("Moneys")
	
	if Title="" or LinkUrl="" or ClassID="" or LastTime="" then
		Call Alert ("����д����","-1")
	end if
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from "&tbname&"_Link where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
	
		rs("Title")				=Title
		rs("ClassID")			=ClassID
		rs("LinkUrl")			=LinkUrl	
		rs("LogoUrl")			=LogoUrl
		rs("AddTime")			=AddTime
		rs("LastTime")			=LastTime
		rs("Num")				=Num
		rs("qq")				=qq
		rs("Content")			=Content
		rs("Moneys")				=Moneys
		
		rs.update
		Call Alert ("�޸ĳɹ�","Admin_Link.asp")
	else
		Call Alert ("�޸Ĵ���","Admin_Link.asp")
	end if
	rs.close
	set rs=nothing
end sub

Sub Admin_ShowClass_Option()
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from "&tbname&"_LinkClass order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">��ѡ�����</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">�������ӷ���</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """" & Rsp("ID") & """" & "")
		 If int(session("YaoLinkClassID"))=Rsp("ID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">|-" & Rsp("LinkName") & "")
	     Response.Write("</option>" ) 
      Rsp.Movenext   
      Loop   
   End if
   Rsp.close:Set Rsp=nothing
End sub

Sub delAll
ID=Trim(Request("ID"))
ytype=Request("ytype")
page=request("page")
If ID="" And Request("Del")<>"�༭" Then
	  Call Alert ("��ѡ���¼!","-1")
ElseIf Request("Del")="����δ��" Then
   set rs=conn.execute("Update "&tbname&"_Link set yn = 0 where ID In(" & ID & ")")
   Response.Write("<script>alert(""�����ɹ���"");location.href=""Admin_Link.asp"";</script>")
ElseIf Request("Del")="�������" Then
   set rs=conn.execute("Update "&tbname&"_Link set yn = 1 where ID In(" & ID & ")")
   Response.Write("<script>alert(""�����ɹ���"");location.href=""Admin_Link.asp"";</script>")
ElseIf Request("Del")="�ƶ�" Then
		If ytype="" then
			Response.Write("<script language=javascript>alert('��ѡ�����!');history.back(1);</script>")
			Response.End
		End if
   set rs=conn.execute("Update "&tbname&"_Link set ClassID = "&ytype&" where ID In(" & ID & ")")
   Response.Write("<script>alert(""�����ɹ���"");location.href=""Admin_Link.asp"";</script>")
ElseIf Request("Del")="ɾ��" Then
	set rs=conn.execute("delete from "&tbname&"_Link where ID In(" & ID & ")")
    Call Alert ("�����ɹ�","Admin_Link.asp") 
ElseIf Request("Del")="�༭" Then
	Server.ScriptTimeout=99999999
	Dim Num,Moneys,AddTimes,LastTime,qq
	For i=1 to request.form("LinkID").count
		LinkID=replace(request.form("LinkID")(i),"'","")
		Num=replace(request.form("Num")(i),"'","")
		Moneys=replace(request.form("Moneys")(i),"'","")
		AddTimes=replace(request.form("AddTimes")(i),"'","")
		LastTime=replace(request.form("LastTime")(i),"'","")
		Title=replace(request.form("Title")(i),"'","")
		LinkUrl=replace(request.form("LinkUrl")(i),"'","")
		qq=replace(request.form("qq")(i),"'","")
		set rs=conn.Execute("select * from "&tbname&"_Link where ID="&LinkID)
		conn.Execute("Update "&tbname&"_Link set Title='"&Title&"',LinkUrl='"&LinkUrl&"',Num="&Num&",Moneys="&Moneys&",AddTime='"&AddTimes&"',LastTime='"&LastTime&"',qq='"&qq&"' where ID="&LinkID)
	next
    Call Alert ("�����ɹ�","Admin_Link.asp")  
End If
End Sub
%>
<!--#include file="Admin_copy.asp"-->
</body>
</html>