<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/md5.asp"-->
<!--#include file="../Inc/Function_Page.asp"-->
<!--#include file="admin_check.asp"-->
<%
Call chkAdmin(11)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="images/Admin_css.css" type=text/css rel=stylesheet>
<script language="javascript" type="text/javascript" src="<%=SitePath%>js/setdate/WdatePicker.js"></script>

</head>

<body>
<table width="95%" border="0" cellspacing="2" cellpadding="3"  align=center class="admintable" style="margin-bottom:5px;">
    <tr><form name="form1" method="get" action="Admin_User.asp">
      <td height="25" bgcolor="f7f7f7">���ٲ��ң�
        <SELECT onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')"  size="1" name="s">
        <OPTION value=""<%If request("s")="" then Response.Write(" selected") end if%>>-=��ѡ��=-</OPTION>
        <OPTION value="?s=all"<%If request("s")="all" then Response.Write(" selected") end if%>>�����û�</OPTION>
        <OPTION value="?s=yn1"<%If request("s")="yn1" then Response.Write(" selected") end if%>>������û�</OPTION>
        <OPTION value="?s=yn0"<%If request("s")="yn0" then Response.Write(" selected") end if%>>δ����û�</OPTION>
        <OPTION value="?s=2"<%If request("s")="2" then Response.Write(" selected") end if%>>24Сʱ��¼�û�</OPTION>
        <OPTION value="?s=1"<%If request("s")="1" then Response.Write(" selected") end if%>>24Сʱע���û�</OPTION>
      </SELECT>      </td>
      <td bgcolor="f7f7f7">
        <input name="keyword" type="text" id="keyword" value="<%=request("keyword")%>">
        <input name="Submit2" type="submit" class="bnt" value="����"></td>
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
	elseif request("action")="updateusergroup" then
		call updateusergroup()
	elseif request("action")="deljf" then
		call deljf()
	elseif request("action")="isyn" then
		call isyn()
	else
		call List()
	end if
 
sub List()
%>
<table border="0" cellspacing="2" cellpadding="3"  align="center" class="admintable">
<tr> 
  <td colspan="9" align=left class="admintitle">�û��б���[<a href="?action=add">����</a>]</td>
</tr>
  <tr align="center"> 
    <td width="15%" class="ButtonList">�û���</td>
    <td width="5%" class="ButtonList">����</td>
    <td width="12%" class="ButtonList">�ȼ�</td>
    <td width="4%" class="ButtonList">�Ա�</td>
    <td width="10%" class="ButtonList">����</td>
    <td width="14%" class="ButtonList">ע��ʱ��</td>
    <td width="13%" class="ButtonList">ע��ɣ�</td>
    <td width="14%" class="ButtonList">�� ��</td>
  </tr>
<%
page=request("page")
Hits=request("hits")
s=Request("s")
Articleclass=request("ClassID")
keyword=request("keyword")
usergroupid=request("usergroupid")
dengji=request("dengji")
Set mypage=new xdownpage
mypage.getconn=conn
mysql="select * from "&tbname&"_User"
	if s="yn0" then
	mysql=mysql&" Where yn=0"
	elseif s="yn1" then
	mysql=mysql&" Where yn=1"
	elseif s="1" then
		If IsSqlDataBase = 1 then
		mysql=mysql&" Where datediff(hh,RegTime,GetDate()) <= 24"
		else
		mysql=mysql&" Where datediff('h',RegTime,Now()) <= 24"
		End if
	elseif s="2" then
		If IsSqlDataBase = 1 then
		mysql=mysql&" Where datediff(hh,LastTime,GetDate()) <= 24"
		else
		mysql=mysql&" Where datediff('h',LastTime,Now()) <= 24"
		End if
	elseif s="vip" then
	mysql=mysql&" Where usergroupid=30"
	elseif keyword<>"" then
	mysql=mysql&" Where UserName like '%"&keyword&"%'"
	elseif usergroupid<>"" then
	mysql=mysql&" Where usergroupid = "&usergroupid&""
	elseif dengji<>"" then
	mysql=mysql&" Where dengji = '"&dengji&"'"
	End if
mysql=mysql&" order by ID desc"
mypage.getsql=mysql
mypage.pagesize=15
set rs=mypage.getrs()
for i=1 to mypage.pagesize
    if not rs.eof then 
%>
    <tr bgcolor="#f1f3f5" onMouseOver="this.style.backgroundColor='#EAFCD5';this.style.color='red'" onMouseOut="this.style.backgroundColor='';this.style.color=''">
    <td height="25" class="tdleft"><%if rs("QQOpenID")<>"" then Echo "<font color=red>��QQ��¼��</font>" end if%><%=rs("UserName")%> (<font color=red> <%=mydb("Select Count([ID]) From ["&tbname&"_Article] Where UserID="&rs("ID")&"",1)(0)%> </font>)</td>
    <td height="25" align="center" class="tdleft"><%=rs("UserMoney")%></td>
    <td height="25" align="center" class="tdleft"><%=UserGroupInfo(rs("usergroupid"),0)%></td>
    <td height="25" align="center"><%If rs("Sex")=1 then Response.Write("��") else Response.Write("Ů") end if%></td>
	<td align="center"><%=rs("province")%><%=rs("city")%></td>
    <td align="center"><%=rs("regtime")%></td>
    <td align="center"><a href="http://www.laoy.net/other/ip.asp?ip=<%=rs("IP")%>" target="_blank" title="�����ѯ��Դ"><u><%=rs("IP")%></u></a></td>
    <td width="11%" align="center">
			<%
			Response.Write "<a href='?action=isyn&yn="&rs("yn")&"&ID=" & rs("ID") & "&page="&request("page")&"'>"
            If rs("yn")=0 Then
               Response.Write "<font color=red>δ��</font>"
            Else
               Response.Write "����"
            End If
            Response.Write "</a>"
           %>|<a href="?action=edit&id=<%=rs("ID")%>">�༭</a>|<a href="?action=del&id=<%=rs("ID")%>&UserName=<%=rs("UserName")%>" onClick="JavaScript:return confirm('ȷ��ɾ�����⽫����ͬ���û�����������һ��ɾ����')">ɾ��</a></td>
  </tr>
<%
        rs.movenext
    else
         exit for
    end if
next
%>
<tr><td colspan=8 class=td>
<div id="page">
	<ul style="text-align:left;">
    <%=mypage.showpage()%>
    </ul>
</div>
</td>
</tr>
</table>
<%
	rs.close
	set rs=nothing
%>
<table border="0" cellspacing="2" cellpadding="3"  align="center" class="admintable">
  <tr>
    <td colspan="2" align=left class="admintitle">����û�����</td>
  </tr>
  <tr>
    <td height="50">
      <form name="form1" method="post" action="?action=deljf">
        <input type="submit" name="button" id="button" value="�û����ֹ���"  onClick="JavaScript:return confirm('ȷ������𣿲��ɻָ���')">
      �����ز������˲��������������û����ֹ��㲢���ɻָ�!
      </form>
    </td>
  </tr>
</table>
<%
end sub

sub add()
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="5" class="admintitle">���ӻ�Ա</th></tr>
<form action="?action=savenew" name="UserReg" method=post>
<tr>
<td width="20%" class=b1_1>��Ա����</td>
<td class=b1_1 colspan=4><input name="UserName" type="text" id="UserName" size="30"></td>
</tr>
<tr> 
<td width="20%" class=b1_1>����</td>
<td colspan=4 class=b1_1><input name="PassWord" type="text" id="PassWord" size="30"></td>
</tr>
<tr>
  <td class=b1_1>�Ա�</td>
  <td colspan=4 class=b1_1><select name="Sex" id="Sex">
    <option value="1">��</option>
    <option value="0">Ů</option>
  </select></td>
</tr>
<tr>
  <td class=b1_1>��������</td>
  <td colspan=4 class=b1_1><input name='Birthday' type='text' class="borderall" onFocus="WdatePicker({isShowClear:false,readOnly:true,startDate:'1960-01-01',minDate:'1960-01-01',maxDate:'1994-12-31',skin:'whyGreen'})" style="width:140px;"/></td>
</tr>
<tr>
  <td class=b1_1>����(ʡ/��)��</td>
  <td colspan=4 class=b1_1><select onChange="setcity();" name='province' style="width:90px;">
    <option value=''>��ѡ��ʡ��</option>
    <option value="�㶫">�㶫</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="�ӱ�">�ӱ�</option>
    <option value="������">������</option>
    <option value="����">����</option>
    <option value="���">���</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="���ɹ�">���ɹ�</option>
    <option value="����">����</option>
    <option value="�ຣ">�ຣ</option>
    <option value="ɽ��">ɽ��</option>
    <option value="�Ϻ�">�Ϻ�</option>
    <option value="ɽ��">ɽ��</option>
    <option value="����">����</option>
    <option value="�Ĵ�">�Ĵ�</option>
    <option value="����">����</option>
    <option value="̨��">̨��</option>
    <option value="���">���</option>
    <option value="�½�">�½�</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="�㽭">�㽭</option>
    <option value="����">����</option>
  </select>
    <select name='city'  style="width:90px;">
    </select>
    <script src="<%=SitePath%>js/getcity.js"></script>
    <script>initprovcity('','');</script>
    <font color="#FF0000">*</font></td>
</tr>
<tr>
  <td class=b1_1>����</td>
  <td colspan=4 class=b1_1><input name="UserMoney" type="text" id="UserMoney" value="0" size="30" maxlength="5"></td>
</tr>
<tr>
  <td class=b1_1>�û�����</td>
  <td colspan=4 class=b1_1><input name="Email" type="text" id="Email" size="30"></td>
</tr>
<tr>
  <td class=b1_1>QQ</td>
  <td colspan=4 class=b1_1><input name="UserQQ" type="text" id="UserQQ" size="30"></td>
</tr>
<tr> 
<td width="20%" class=b1_1></td>
<td colspan=4 class=b1_1><input name="Submit" type="submit" class="bnt" value="�� ��"></td>
</tr></form>
</table>
<%
end sub

sub savenew()
	UserName=trim(request.form("UserName"))
	PassWord=trim(request.form("PassWord"))
	Email=trim(request.form("Email"))
	sex=trim(request.form("sex"))
	UserQQ=trim(request.form("UserQQ"))
	UserMoney=trim(request.form("UserMoney"))
	Birthday = 		request.form("Birthday")
	Province = 		request.form("Province")
	City = 			request.form("City")
	
	if UserName="" or PassWord="" then
		Call Alert ("�û��������벻��Ϊ��",-1)
	end if
	if Birthday="" then
		Call Alert ("���ղ���Ϊ��",-1)
	end if

	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from "&tbname&"_User where UserName='"&UserName&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.AddNew 
		rs("UserName")		=UserName
		rs("Email")			=Email
		rs("UserQQ")		=UserQQ
		rs("Sex")			=sex
		rs("PassWord")		=md5(PassWord,16)
		rs("UserMoney")		=UserMoney
		rs("Birthday")		=Birthday
	  	rs("usergroupid")	=1
	  	rs("yn")=1
		rs("RegTime")=Now
		rs("IP")=GetIP
		rs("province")=Province
		rs("city")=City

		rs.update
		Call Alert ("���ӳɹ���","Admin_User.asp")
	else
		Call Alert ("���û��Ѵ��ڣ�",-1)
	end if
	rs.close
	set rs=nothing
end sub

sub del()
	id=request("id")
	set rs=conn.execute("delete from "&tbname&"_User where id="&id)
	set rs=conn.execute("delete from "&tbname&"_Article where UserID="&id)
	Call Alert ("ɾ���ɹ�","Admin_User.asp")
end sub

sub isyn()
	id=request("id")
	yn=request("yn")
	page=request("page")
	If yn=1 then
	yn=0
	else
	yn=1
	End if
	set rs=conn.execute("Update ["&tbname&"_User] set yn="&yn&" Where ID="&id)
	Response.Redirect "Admin_User.asp?page="&page&""
end sub

sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from "&tbname&"_User where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form action="?action=savedit" name="UserReg" method=post>
<tr> 
    <td colspan="5" class="admintitle">�޸Ļ�Ա</td>
</tr>
<tr> 
<td width="20%" class="b1_1">��Ա��¼��</td>
<td colspan=4 class=b1_1><%=rs("UserName")%></td>
</tr>
<tr>
  <td class="b1_1">�û���</td>
  <td colspan=4 class=b1_1><select size=1 name="usergroupid">
<%
set trs=conn.Execute("select UserGroupID,GroupName from "&tbname&"_UserGroup order by UserPowers asc")
do while not trs.eof
response.write "<option value="&trs(0)
if rs("usergroupid")=trs(0) then response.write " selected "
response.write ">"&tRs(1)&"</option>" & VbCrLf
trs.movenext
loop
trs.close
set trs=nothing
%>
</select></td>
</tr>
<tr>
  <td class="b1_1">����</td>
  <td colspan=4 class=b1_1><input name="PassWord" type="text" id="PassWord" size="30" maxlength="12">
    *���޸�������,ԭ����:<%=rs("PassWord")%></td>
</tr>
<tr>
  <td class="b1_1">�Ա�</td>
  <td colspan=4 class=b1_1><select name="Sex" id="Sex">
    <option value="1"<%If rs("Sex")=1 then Response.Write(" selected") End if%>>��</option>
    <option value="0"<%If rs("Sex")=0 then Response.Write(" selected") End if%>>Ů</option>
  </select></td>
</tr>
<tr>
  <td class="b1_1">����(ʡ/��)��</td>
  <td colspan=4 class=b1_1><select name='province' id="province" style="width:90px;" onChange="setcity();">
    <option value=''>��ѡ��ʡ��</option>
    <option value="�㶫">�㶫</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="�ӱ�">�ӱ�</option>
    <option value="������">������</option>
    <option value="����">����</option>
    <option value="���">���</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="���ɹ�">���ɹ�</option>
    <option value="����">����</option>
    <option value="�ຣ">�ຣ</option>
    <option value="ɽ��">ɽ��</option>
    <option value="�Ϻ�">�Ϻ�</option>
    <option value="ɽ��">ɽ��</option>
    <option value="����">����</option>
    <option value="�Ĵ�">�Ĵ�</option>
    <option value="����">����</option>
    <option value="̨��">̨��</option>
    <option value="���">���</option>
    <option value="�½�">�½�</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="�㽭">�㽭</option>
    <option value="����">����</option>
  </select>
    <select name='city' id="city"  style="width:90px;">
    </select>
    <script src="<%=SitePath%>js/getcity.js"></script>
    <script>initprovcity('<%=rs("province")%>','<%=rs("city")%>');</script>
    <font color="#FF0000">*</font></td>
</tr>
<tr>
  <td class="b1_1">����</td>
  <td colspan=4 class=b1_1><input name="UserMoney" type="text" id="UserMoney" value="<%=rs("UserMoney")%>" size="30"></td>
</tr>
<tr>
  <td class="b1_1">�û�����</td>
  <td colspan=4 class=b1_1><input name="Email" type="text" id="Email" value="<%=rs("Email")%>" size="30"></td>
</tr>
<tr>
  <td class=b1_1>����</td>
  <td colspan=4 class=b1_1><input name="Birthday" type="text" id="Birthday" value="<%=rs("Birthday")%>" size="30"></td>
</tr>
<tr>
  <td class=b1_1>QQ</td>
  <td colspan=4 class=b1_1><input name="UserQQ" type="text" id="UserQQ" value="<%=rs("UserQQ")%>" size="30"></td>
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
	Dim id
	id=request.form("id")
	'UserName=trim(request.form("UserName"))
	PassWord=trim(request.form("PassWord"))
	Email=trim(request.form("Email"))
	Birthday=trim(request.form("Birthday"))
	UserQQ=trim(request.form("UserQQ"))
	usergroupid=trim(request.form("usergroupid"))
	Sex=trim(request.form("Sex"))
	Province = 		request.form("Province")
	City = 			request.form("City")
	UserMoney=		request.form("UserMoney")
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from "&tbname&"_User where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
		'rs("UserName")		= UserName
		rs("Email")			= Email
		If birthday<>"" then
		rs("Birthday")		=Birthday
		end if
		rs("UserQQ")		=UserQQ
		rs("Sex")			=Sex
		If PassWord<>"" then
		rs("PassWord")		=md5(PassWord,16)
		end if
		rs("usergroupid")	=usergroupid
		rs("province")		=Province
		rs("city")			=City
		rs("UserMoney")		=UserMoney
		
		rs.update
		Response.write"<script>alert(""��ϲ,�޸ĳɹ���"");location.href=""Admin_User.asp"";</script>"
	else
		Response.write"<script>alert(""�޸Ĵ���"");location.href=""Admin_User.asp"";</script>"
	end if
	rs.close
	set rs=nothing
end sub

sub deljf()
	set rs=conn.execute("Update "&tbname&"_User set UserMoney = 0")
	Call Alert ("�����ѹ��㣡","admin_user.asp")
end sub
%>
<!--#include file="Admin_copy.asp"-->
</body>
</html>