<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/saveimage.asp"-->
<!--#include file="../Inc/Cls_vbsPage.asp"-->
<!--#include file="Admin_check.asp"--><html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equtv="Content-Type" content="text/html; charset=gb2312" />
<title>��վ��̨����</title>
<link href="images/Admin_css.css" type=text/css rel=stylesheet>
<script src="js/admin.js"></script>
<script language="JavaScript" src="../js/Jquery.js"></script>
<script charset="utf-8" src="../KindEditor/kindeditor.js"></script>
<script charset="utf-8" src="../KindEditor/lang/zh_CN.js"></script>
<script>
	KindEditor.ready(function(KE) {
		window.editor = KE.create('#Content', {
			minWidth : 700,minHeight:300,allowFlashUpload:false,allowMediaUpload:false,allowFileUpload:false,
			cssPath : '../KindEditor/plugins/code/prettify.css',
			uploadJson : '<%=sitepath&SiteAdmin%>/upload.asp?action=simgedit',
			afterBlur: function(){this.sync();}
		});
	});
</script>
</head>

<body>
<script language=javascript>
function CheckForm()
{ 
  if (document.myform.Title.value==""){
	alert("����д���⣡");
	document.myform.Title.focus();
	return false;
  }
  if (document.myform.ClassID.value==""){
	alert("��ѡ����࣡");
	document.myform.ClassID.focus();
	return false;
  }
  if (document.myform.Content.value==""){
	alert("����д���ݣ�");
	return false;
  }
  if (document.myform.Hits.value==""){
	alert("����д���������");
	document.myform.Hits.focus();
	return false;
  }
  var filter=/^\s*[0-9]{1,6}\s*$/;
  if (!filter.test(document.myform.Hits.value)) { 
	alert("���������д����ȷ,��������д��"); 
	document.myform.Hits.focus();
	return false; 
  }
}
			function ResumeError()
			{return true;}
			window.onerror = ResumeError;
			$(document).ready(function(){
				$(parent.frames["main"].document).find("#Button1").attr("disabled",false);
			  $('#KeyLinkByTitle').click(function(){
			    GetKeyTags();
			  });
			});
			function GetKeyTags()
			{
			  var text=escape($('input[name=Title]').val());
			  if (text!=''){
				  $('#KeyWord').val('�Ե�,��y�����Զ���ȡ...');
				  $("#KeyWord").attr("disabled",true);
				  $.get("ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWord').val(unescape(data));
					$('#KeyWord').attr("disabled",false);
				  });
			  }else{
			   alert('�����������!');
			  }
			}
</script>
<table width="95%" border="0" cellspacing="2" cellpadding="3"  align=center class="admintable" style="margin-bottom:5px;">
    <tr><form name="form1" method="get" action="Admin_Article.asp?page=<%=page%>">
      <td height="25" bgcolor="f7f7f7">���ٲ��ң�
        <SELECT onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')"  size="1" name="s">
        <OPTION value="" selected>-=��ѡ��=-</OPTION>
        <OPTION value="?s=all">��������</OPTION>
        <OPTION value="?s=yn0">���������</OPTION>
        <OPTION value="?s=yn1">δ�������</OPTION>
        <OPTION value="?s=yn2">��Ա˽������</OPTION>
        <OPTION value="?s=istop">�̶�����</OPTION>
        <OPTION value="?s=ishot">�Ƽ�����</OPTION>
        <OPTION value="?s=img">������ͼ����</OPTION>
        <OPTION value="?s=url">ת����������</OPTION>
        <OPTION value="?s=user">��Ա���������</OPTION>
        <OPTION value="?p=hits">�������������</OPTION>
        <OPTION value="?p=dateandtime">������ʱ������</OPTION>
      </SELECT>      </td>
      <td align="center" bgcolor="f7f7f7">
        <a href="?hits=1"></a>
        <input name="keyword" type="text" id="keyword" value="<%=request("keyword")%>" class="s26">
        <input type="submit" name="Submit2" class="bnt" value="����">
      </form></td>
      <td align="right" bgcolor="f7f7f7">��ת����
        <select name="ClassID" id="ClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
    <%
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from "&tbname&"_Class Where TopID = 0 and link=0 order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">��ѡ�����</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">������ӷ���</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """?ClassID=" & Rsp("ID") & """" & "")
		 If int(request("ClassID"))=Rsp("ID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">|-" & Rsp("ClassName") & " ("&Mydb("Select Count([ID]) From ["&tbname&"_Article] Where ClassID="&rsp("ID")&"",1)(0)&")</option>")  & VbCrLf
		 
		    Sqlpp ="select * from "&tbname&"_Class Where TopID="&Rsp("ID")&" and link=0 order by num"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof 
				Response.Write("<option value=" & """?ClassID=" & Rspp("ID") & """" & "")
				If int(request("ClassID"))=Rspp("ID") then
				Response.Write(" selected" ) 
				End if
         		Response.Write(">��|-" & Rspp("ClassName") & " ("&Mydb("Select Count([ID]) From ["&tbname&"_Article] Where ClassID="&rspp("ID")&"",1)(0)&")")
				Response.Write("</option>" )   & VbCrLf
			Rspp.Movenext   
      		Loop
			rspp.close
			set rspp=nothing
      Rsp.Movenext   
      Loop   
   End if
	rsp.close
	set rsp=nothing
	%>
        </select></td>
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
<form name="myform" id="myform" method="POST" action="Admin_Article.asp?action=delAll">
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="1" bgcolor="#F2F9E8" class="admintable">
<tr> 
  <td colspan="5" align=left class="admintitle">�����б�[<a href="?action=add">���</a>]</td></tr>
    <tr bgcolor="#f1f3f5" style="font-weight:bold;">
    <td width="5%" align="center" class="ButtonList">&nbsp;</td>
    <td width="53%" align="center" class="ButtonList">��������</td>
    <td width="20%" align="center" class="ButtonList">����ʱ��</td>
    <td width="7%" align="center" class="ButtonList">���</td>
    <td width="15%" align="center" class="ButtonList">����</td>    
    </tr>
<%
Hits=request("hits")
s=Request("s")
p=Request("p")
istitle=Request("istitle")
Articleclass=request("ClassID")
keyword=request("keyword")
Dim ors
Set ors=new Cls_vbsPage	'��������
Set ors.Conn=conn		'�õ����ݿ����Ӷ���
With ors
	.PageSize=15		'ÿҳ��¼����
	.PageName="page"	'cookies����
	.DbType=""
	.RecType=-1
	'��¼����(>0Ϊ����ȡֵ�ٸ�����߹̶�ֵ,0ִ��count���ô�cookies,-1ִ��count������cookies)
	.JsUrl=""			'Cls_jsPage.js��·��
	.Pkey="ID"			'����
	.Field="ID,Title,Author,ClassID,DateAndTime,Hits,Yn,IsTop,IsHot,IsFlash,Images,UserID"
	.Table=tbname&"_Article"
	
If yaoadmintype<>1 and (yaoManagePower<>"" or yaoWritePower<>"") then
	w="ClassID in ("&yaoManagePower&") and "
Else
	w="yn<>3 and "
End if
	
	if s="yn0" then
	.Condition=w&"yn=0"
	elseif s="yn1" then
	.Condition=w&"yn=1"
	elseif s="yn2" then
	.Condition=w&"yn=2"
	elseif s="istop" then
	.Condition=w&"istop=1"
	elseif s="k" then
	.Condition=w&"title is null"
	elseif s="ishot" then
	.Condition=w&"ishot=1"
	elseif s="img" then
	.Condition=w&w&"images<>''"
	elseif s="url" then
	.Condition=w&"LinkUrl<>''"
	elseif s="user" then
	.Condition=w&"UserID>0"
	elseif Articleclass<>"" then
	.Condition=w&"ClassID="&Articleclass&""
	'elseif keyword<>"" and istitle<>1 then
	'.Condition=w&"(title like '%"&KeyWord&"%' or content like '%"&KeyWord&"%' or ID like '%"&KeyWord&"%')"
	elseif keyword<>"" then
	.Condition=w&"(title like '%"&KeyWord&"%' or ID like '%"&KeyWord&"%')"
	else
	.Condition=w&"1=1"
	End if
	If p<>"" then
	.OrderBy=p&" Desc"
	else
	.OrderBy="ID Desc"
	End if
End With
on error resume next
RecCount=ors.RecCount()'��¼����
Rs=ors.ResultSet()		'����ResultSet
NoI=0
If  RecCount<1 Then%>
<tr bgcolor="#ffffff">
    <td height="25" colspan="5" style="color:#F00;">û�м�¼</td>
    </tr>
<%
Else 
	iPageCount=Abs(Int(-Abs(RecCount/artlistnum)))
    For i=0 To Ubound(Rs,2)
	NoI=NoI+1
%>
    <tr bgcolor="#ffffff" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
    <td align="CENTER"><input type="checkbox" value="<%=rs(0,i)%>" name="ID" onClick="unselectall(this.form)" style="border:0;"></td>
    <td align="left">[<a href="?ClassID=<%=rs(3,i)%>"><%=getclass(rs(3,i),"classname")%></a>]<a href="<%=apath(rs(0,i),0)%>" target="_blank"><%=rs(1,i)%></a> <%if rs(7,i)=1 then Response.Write("<font color=red>[��]</font>") end if%><%if rs(8,i)=1 then Response.Write("<font color=green>[��]</font>") end if%><%if rs(10,i)<>"" then Response.Write("<font color=blue>[ͼ]</font>") end if%><%if rs(9,i)=1 then Response.Write("<font color=orange>[��]</font>") end if%><%If rs(11,i)<>0 then Response.Write(" <font color=blue>["&UserInfo(rs(11,i),2)&"]</font>") end if%></td>
    <td align="center"><%=rs(4,i)%></td>
    <td align="center"><%=rs(5,i)%></td>
    <td align="center"><%If rs(6,i)=0 then Response.Write("����") end if:If rs(6,i)=1 then Response.Write("<font color=red>δ��</font>") end if:If rs(6,i)=2 then Response.Write("<font color=blue>˽��</font>") end if%>|<a href="?action=edit&id=<%=rs(0,i)%>&page=<%=page%>">�༭</a>|<%If Mydb("Select Count([ID]) From ["&tbname&"_Pl] Where ArticleID="&rs(0,i)&"",1)(0)>0 Then Response.Write("<a href=""Admin_pl.asp?id="&rs(0,i)&""">����</a>") else Response.Write("<a href=""#""><font style='color:#ccc;'>����</font></a>") End if%></td>    
    </tr>
<%
    Next	
End If
%>
<tr><td align="center" bgcolor="f7f7f7"><input name="Action" type="hidden"  value="Del"><input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox" style="border:0"></td>
  <td colspan="4" bgcolor="f7f7f7" align="left"><font color=red>�ƶ�����</font>
    <select id="ytype" name="ytype">
      <%call Admin_ShowClass_Option()%>
    </select>
    &nbsp;
    <input name="Del" type="submit" class="bnt" id="Del" value="�ƶ�">
	<input name="Del" type="submit" class="bnt" id="Del" value="����ʱ��">
    <input name="Del" type="submit" class="bnt" id="Del" value="ɾ��">
    <input name="Del" type="submit" class="bnt" id="Del" value="����δ��">
    <input name="Del" type="submit" class="bnt" id="Del" value="�������">
    <input name="Del" type="submit" class="bnt" id="Del" value="�Ƽ�">
    <input name="Del" type="submit" class="bnt" id="Del" value="����Ƽ�">
    <input name="Del" type="submit" class="bnt" id="Del" value="�̶�">
    <input name="Del" type="submit" class="bnt" id="Del" value="����̶�"></td>
  </tr><tr><td bgcolor="f7f7f7" colspan="5">
<div id="page">
	<ul style="text-align:left;">
    <%ors.ShowPage()%>
    </ul>
</div>
</td></tr></table>
</form>
<%
	rs.close
	set rs=nothing
end sub

sub add()
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form onSubmit="return CheckForm();" action="?action=savenew" id="myform" name="myform" method=post>
<tr> 
    <td colspan="2" align=left class="admintitle">�������</td></tr>
<tr> 
<td width="20%" class="b1_1">����</td>
<td class="b1_1"><input name="Title" type="text" id="Title" size="40" maxlength="50" onBlur="CheckName();" onChange="CheckName();">
		<input name="TitleFontColor" type="text" class="color {adjust:false,hash:true}" id="TitleFontColor" size="10" maxlength="7"> &nbsp;<span class="note">ע�����50���ַ�</span></td>
</tr>
<tr>
  <td class="b1_1">�ؼ��֡�<a href="#" id="KeyLinkByTitle" style="color:green">���ݱ����Զ���ȡ</a>��</td>
  <td class="b1_1"><input name="KeyWord" type="text" id="KeyWord" size="40" maxlength="50">&nbsp;<span class="note">����ؼ�����|����</span></td>
</tr>
<tr>
  <td class="b1_1">����</td>
  <td class="b1_1">
    <input name="Author" type="text" id="Author" value="����" size="40" maxlength="200">    </td>
</tr>
<tr>
  <td class="b1_1">��Դ</td>
  <td class="b1_1"><span class="td">
    <input name="CopyFrom" type="text" id="CopyFrom" value="����" size="40" maxlength="200">
  </span></td>
</tr>
<tr>
  <td class="b1_1">����</td>
  <td class="b1_1"><select ID="ClassID" name="ClassID">
    <%call Admin_ShowClass_Option()%></select>&nbsp;<span class="note">����ж�������,��ѡ���������.</span></td>
</tr>
<tr>
  <td class="b1_1">ת������</td>
  <td class="b1_1"><input name="LinkUrl" type="text" size="50" disabled>
		  <INPUT name="UseLinkUrl" type="checkbox" class="noborder" id="UseLinkUrl" onClick="if(this.checked){
			  LinkUrl.disabled = false;
		  }else{
			  LinkUrl.disabled = true;
		  }" value="Yes">
      ʹ��ת������		</td>
</tr>
<tr>
  <td class="b1_1">�������</td>
  <td class="b1_1"><input name="Hits" type="text" id="Hits" value="0" size="6" maxlength="10"></td>
</tr>
<tr>
  <td rowspan="2" class="b1_1">����ͼ</td>
  <td class="b1_1"><input name="Images" type="text" id="Images" size="50" maxlength="200"></td>
</tr>
<tr>
  <td class="b1_1"><iframe src="upload.asp?action=simg" width="400" height="25" frameborder="0" scrolling="no"></iframe></td>
</tr>
<tr>
  <td class="b1_1">����ժҪ<br><span class="note">����д���Զ���ȡ���Ŀ�ʼ�����ַ�</span></td>
  <td class="b1_1"><textarea name="ArtDescription" id="ArtDescription" cols="60" rows="6"></textarea></td>
</tr>
<tr>
  <td valign="top" class="b1_1"><p>����</p>
    <p>����ʱ��<br>
      <input name="DateAndTime" type="text" id="DateAndTime" value="<%=FormatDate(NOW,1,0)%>">
    </p>
    <p>
      <input name="sSaveFileSelect" type="checkbox" class="noborder" id="sSaveFileSelect" value="True">
    ���������е�ͼƬ������</p><br><br><p>�����ϴ���</p><br><p><iframe src="upload.asp?action=fileup" width="100%" height="40" frameborder="0" scrolling="no"></iframe></p></td>
  <td class="b1_1"><textarea name="Content" style="width:100%;height:350px;" id="Content"></textarea></td>
</tr>
<tr>
  <td class="b1_1">����ѡ��</td>
  <td class="b1_1">�̶�
    <input name="IsTop" type="checkbox" class="noborder" id="IsTop" value="1">
    �Ƽ�
    <input name="IsHot" type="checkbox" class="noborder" id="IsHot" value="1">
    �õ�
    <input name="IsFlash" type="checkbox" class="noborder" id="IsFlash" value="1"></td>
</tr>
<tr>
  <td class="b1_1">�Զ���ҳ<strong>����</strong></td>
  <td class="b1_1"><input name="PageNum" type="text" id="PageNum" value="0" size="6" maxlength="4">
    <span class="note">��ע:����������м������ֶ���ҳ��,����д0</span></td>
</tr>
<tr>
  <td class="b1_1">���Ȩ��</td>
  <td class="b1_1"><select name="ReadPower" size=4 multiple id="ReadPower">
    <%=ShowlevelOption(-1)%>
  </select>    
  <span class="note">ע����ѡ���а����ο͵ȼ�ʱ�������û����ɲ鿴���ɰ�סCtrl��ѡ!��ѡ��̳���ĿȨ�ޡ�</span></td>
</tr>
<%If Mydb("Select Count([ID]) From ["&tbname&"_Vote]",1)(0)>0 then%>
<tr>
  <td class="b1_1">������ʾͶƱ</td>
  <td class="b1_1"><select name="Vote" size=4 multiple id="Vote">
    <%=ShowVoteList(0)%>
  </select><span class="note">���ɰ�סCtrl��ѡ</span></td>
</tr><%End if%>
<tr> 
<td width="20%" class="b1_1"></td>
<td class="b1_1"><input name="button" type="submit" class="bnt" value="�� ��"></td>
</tr>
</form>
</table>
<%
end sub

sub savenew()
	Title			=Replace(trim(request.form("Title")),"'","")
	Hits			=trim(request.form("Hits"))
	ClassID			=trim(request.form("ClassID"))
	Content			=request.form("Content")
	Author			=trim(request.form("Author"))
	CopyFrom		=trim(request.form("CopyFrom"))
	KeyWord			=trim(request.form("KeyWord"))
	IsTop			=request.form("IsTop")
	IsHot			=request.form("IsHot")
	IsFlash			=request.form("IsFlash")
	ArtDescription	=request.form("ArtDescription")
	TitleFontColor	=request.form("TitleFontColor")
	Images			=trim(request.form("Images"))
	PageNum			=trim(request.form("PageNum"))
	LinkUrl			=trim(request.form("LinkUrl"))
	UseLinkUrl		=request.form("UseLinkUrl")
	DateAndTime		=trim(request.form("DateAndTime"))
	sSaveFileSelect =request.Form("sSaveFileSelect")
	Vote			=request.form("Vote")
	Vote		 	=replace(Vote," ","")
	ReadPower		=request.form("ReadPower")
	ReadPower 		=Replace(ReadPower," ","")
	if Title="" or ClassID="" then
		Call Alert ("����д�������ύ",-1)
	end if
	If ClassID="-1" then
		Call Alert ("��ѡ����ȷ�ķ��࣡",-1)
	End if		
	If LinkUrl="" and Content="" then
		Call Alert ("��û����д����",-1)
	End if

	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from "&tbname&"_Article where Title='"&Title&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.AddNew 
		rs("Title")				=Title
		rs("Hits")				=Hits
		rs("ClassID")			=ClassID
		rs("LinkUrl")			=LinkUrl
		rs("Author")			=Author
		rs("CopyFrom")			=CopyFrom
		rs("KeyWord")			=KeyWord
		If IsTop=1 then
			rs("IsTop")				=1
		else
			rs("IsTop")				=0
		end if
		If IsHot=1 then
			rs("IsHot")				=1
		else
			rs("IsHot")				=0
		end if
		If IsFlash=1 then
			rs("IsFlash")			=1
		else
			rs("IsFlash")			=0
		end if
		If ArtDescription<>"" then
			rs("ArtDescription")		=ArtDescription
		Else
			rs("ArtDescription")		=Left(LoseHtml(Content),100)
		End if
		rs("TitleFontColor")	=TitleFontColor
		If Images<>"" then
		rs("Images")			=Images
		Else
		  rs("Images")=GetImg(Content)
		  Cimg=""
		  Cimg=GetImg(Content)
		'   If Cimg<>"" then
		'   rs("IsFlash")=1
		'   Else
		'   rs("IsFlash")=0
		'   End if
		End if
		rs("yn")				=0
		rs("PageNum")			=PageNum
		rs("DateAndTime")		=DateAndTime
		rs("Vote")				=Vote
		rs("ReadPower")			=ReadPower
		rs("UserName")			=myadminuser

		If sSaveFileSelect="True" Then
      		Rs("Content")=ReplaceRemoteUrl(Content,SitePath&SiteUp&"/" & Year(Now) & right("0" & Month(Now), 2) & "/",sFileExt)
     	Else
      		Rs("Content")=Content
		End If

		rs.update
		session("YaoClassID")=ClassID
		Call Alert ("��ӳɹ���","Admin_Article.asp")
	else
		Call Alert ("�������Ѵ��ڣ�",-1)
	end if
	rs.close
	set rs=nothing
end sub

sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from "&tbname&"_Article where id="& id &""
rs.open sql,conn,1,1

Content=server.htmlencode(rs("Content"))
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form onSubmit="return CheckForm();" id="myform" name="myform" action="?action=savedit" method=post>
<tr> 
    <td colspan="5" class="admintitle">�޸�����</td></tr>
<tr>
  <td width="20%" bgcolor="#FFFFFF" class="b1_1">����</td>
  <td colspan=4 class=b1_1><input name="Title" type="text" value="<%=rs("Title")%>" size="40" maxlength="50">
		<input name="TitleFontColor" type="text" class="color {adjust:false,hash:true}" id="TitleFontColor" value="<%=rs("TitleFontColor")%>" size="10" maxlength="7">&nbsp;<span class="note">ע�����50���ַ�</span></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">�ؼ��֡�<a href="#" id="KeyLinkByTitle" style="color:green">���ݱ����Զ���ȡ</a>��</td>
  <td colspan=4 class=b1_1><input name="KeyWord" type="text" id="KeyWord" value="<%=rs("KeyWord")%>" size="40">&nbsp;<span class="note">����ؼ�����|����</span></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">����</td>
  <td colspan=4 class=b1_1><input name="Author" type="text" id="Author" value="<%=rs("Author")%>" size="40"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">¼��Ա</td>
  <td colspan=4 class=b1_1><input name="UserName" type="text" id="UserName" value="<%=rs("UserName")%>" size="40"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">��Դ</td>
  <td colspan=4 class=b1_1><input name="CopyFrom" type="text" id="CopyFrom" value="<%=rs("CopyFrom")%>" size="40"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">����</td>
  <td colspan=4 class=b1_1><select ID="ClassID" name="ClassID">
   <%
   Set Rsp=server.CreateObject("adodb.recordset") 
   Sqlp ="select * from "&tbname&"_Class Where TopID = 0 and link=0 order by num"   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">��ѡ�����</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">������ӷ���</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option")
         If rs("ClassID")=Rsp("ID") Then
            Response.Write(" selected")
         End If
		 If Yao_MyID(rsp("ID"))<>"0" then
		 		Response.Write(" value=""-1"" style='background:#f7f7f7;color:#ccc;'")
		 Else
		 	If yaoadmintype<>1 then
				If instr(","&yaoWritePower&",",","& rsp("ID") &",") <> 0 or instr(","&yaoManagePower&",",","& rsp("ID") &",") <> 0 then
					Response.Write(" value=" & """" & Rsp("ID") & """" & " style='color:#0000ff;'")
				Else
					Response.Write(" value=""-1"" style='background:#f7f7f7;color:#ccc;'")
				End if
			else
				Response.Write(" value=" & """" & Rsp("ID") & """" & " style='color:#0000ff;'")
			End if
		 End if
         Response.Write(">|-" & Rsp("ClassName") & "</option>")  & VbCrLf
		 
		 Sqlpp ="select * from "&tbname&"_Class Where TopID="&Rsp("ID")&" and link=0 order by num"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof 
				Response.Write("<option")
				If rs("ClassID")=Rspp("ID") Then
            	Response.Write(" selected")
         		End If
				If yaoadmintype<>1 then
					If instr(","&yaoWritePower&",",","& rspp("ID") &",") <> 0 or instr(","&yaoManagePower&",",","& rspp("ID") &",") <> 0 then
						Response.Write(" value=" & """" & Rspp("ID") & """" & " style='color:#0000ff;'")
					Else
						Response.Write(" value=""-1"" style='background:#f7f7f7;color:#ccc;'")
					End if
				else
					Response.Write(" value=" & """" & Rspp("ID") & """" & " style='color:#0000ff;'")
				End if
         		Response.Write(">��|-" & Rspp("ClassName") & "")
				Response.Write("</option>" ) 
			Rspp.Movenext   
      		Loop
			rspp.close
			set rspp=nothing
      Rsp.Movenext   
      Loop   
   End if
   rsp.close
   set rsp=nothing
   %>
  </select>  </td></tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">ת������</td>
  <td colspan=4 class=b1_1><input name="LinkUrl" type="text" size="50" value="<%=rs("Linkurl")%>" disabled>
		  <INPUT name="UseLinkUrl" type="checkbox" class="noborder" id="UseLinkUrl" onClick="if(this.checked){
			  LinkUrl.disabled = false;
		  }else{
			  LinkUrl.disabled = true;
		  }" value="Yes"<%if rs("LinkUrl")<>"" then Response.Write(" checked") end if%>>
      ʹ��ת������		</td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">�������</td>
  <td colspan=4 class=b1_1><input name="Hits" type="text" id="Hits" value="<%=rs("Hits")%>" size="6" maxlength="5"></td>
</tr>
<tr>
  <td rowspan="2" bgcolor="#FFFFFF" class="b1_1">����ͼ</td>
  <td colspan=4 class=b1_1><input name="Images" type="text" id="Images" value="<%=rs("Images")%>" size="50" maxlength="200"></td></tr>
<tr>
  <td colspan=4 class=b1_1><iframe src="upload.asp?action=simg" width="400" height="25" frameborder="0" scrolling="no"></iframe></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">����ժҪ</td>
  <td colspan=4 class=b1_1><textarea name="ArtDescription" id="ArtDescription" cols="60" rows="6"><%=rs("ArtDescription")%></textarea></td></tr>
<tr>
  <td valign="top" bgcolor="#FFFFFF" class="b1_1"><p>����</p>
    <p>����ʱ��<br>
      <input name="DateAndTime" type="text" id="DateAndTime" value="<%=FormatDate(rs("DateAndTime"),1,0)%>">
</p>
    <p>
      <input name="sSaveFileSelect" type="checkbox" class="noborder" id="sSaveFileSelect" value="True">
���������е�ͼƬ������</p><br><br><p>�����ϴ���</p><br><p><iframe src="upload.asp?action=fileup" width="100%" height="40" frameborder="0" scrolling="no"></iframe></p></td>
  <td colspan=4 class=b1_1><textarea name="Content" style="width:100%;height:350px;" id="Content"><%=Content%></textarea></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">����ѡ��</td>
  <td colspan=4 class=b1_1>�̶�
    <input name="IsTop" type="checkbox" class="noborder" id="IsTop" value="1"<%if rs("IsTop")=1 then Response.Write("checked") end if%>>
�Ƽ�
<input name="IsHot" type="checkbox" class="noborder" id="IsHot" value="1"<%if rs("IsHot")=1 then Response.Write("checked") end if%>>
�õ�
<input name="IsFlash" type="checkbox" class="noborder" id="IsFlash" value="1"<%if rs("IsFlash")=1 then Response.Write("checked") end if%>></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">�Զ���ҳ<strong>����</strong></td>
  <td colspan=4 class=b1_1><input name="PageNum" type="text" id="PageNum" value="<%=rs("PageNum")%>" size="6" maxlength="4"><span class="note">��ע:����������м������ֶ���ҳ��,����д0</span></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">���Ȩ��</td>
  <td colspan=4 class=b1_1><select name="ReadPower" size=4 multiple id="ReadPower">
    <%=ShowlevelOption(rs("ReadPower"))%>
  </select>    <span class="note">ע����ѡ���а����ο͵ȼ�ʱ�������û����ɲ鿴���ɰ�סCtrl��ѡ!��ѡ��̳���ĿȨ�ޡ�</span></td>
</tr>
<%If Mydb("Select Count([ID]) From ["&tbname&"_Vote]",1)(0)>0 then%>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">������ʾͶƱ</td>
  <td colspan=4 class=b1_1><select name="Vote" size=4 multiple id="Vote">
    <%=ShowVoteList(rs("Vote"))%>
  </select>
    <span class="note">���ɰ�סCtrl��ѡ</span></td>
</tr><%End if%>
<tr> 
<td width="20%" class="b1_1"></td>
<td colspan=4 class=b1_1><input name="id" type="hidden" value="<%=rs("ID")%>"><input name="page" type="hidden" value="<%=page%>"><input name="button" type="submit" class="bnt" value="�� ��"></td>
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
	page=request.form("page")
	Title			=trim(request.form("Title"))
	Hits			=trim(request.form("Hits"))
	ClassID			=trim(request.form("ClassID"))
	Content			=request.form("Content")
	Author			=trim(request.form("Author"))
	UserName		=trim(request.form("UserName"))
	CopyFrom		=trim(request.form("CopyFrom"))
	KeyWord			=trim(request.form("KeyWord"))
	IsTop			=request.form("IsTop")
	IsHot			=request.form("IsHot")
	IsFlash			=request.form("IsFlash")
	ArtDescription	=request.form("ArtDescription")
	TitleFontColor	=request.form("TitleFontColor")
	Images			=trim(request.form("Images"))
	PageNum			=trim(request.form("PageNum"))
	LinkUrl			=trim(request.form("LinkUrl"))
	UseLinkUrl		=request.form("UseLinkUrl")
	DateAndTime		=trim(request.form("DateAndTime"))
	sSaveFileSelect =request.Form("sSaveFileSelect")
	Vote			=request.form("Vote")
	Vote		 	=replace(Vote," ","")
	ReadPower		=request.form("ReadPower")
	ReadPower 		=Replace(ReadPower," ","")
	
	if Title="" or ClassID="" then
		Call Alert ("����д�������ύ","-1")
	end if
	If ClassID="-1" then
		Call Alert ("��ѡ����ȷ�ķ��࣡",-1)
	End if		
	If LinkUrl="" and Content="" then
		Call Alert ("��û����д����","-1")
	End if
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from "&tbname&"_Article where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
	
		rs("Title")				=Replace(Title,CHR(34),"&quot;")
		rs("Hits")				=Hits
		rs("ClassID")			=ClassID
		rs("LinkUrl")			=LinkUrl
		rs("Author")			=Author
		If UserName<>"" then
		rs("UserName")			=UserName
		Else
		rs("UserName")			=myadminuser
		End if
		rs("CopyFrom")			=CopyFrom
		rs("KeyWord")			=KeyWord
		If IsTop=1 then
			rs("IsTop")				=1
		else
			rs("IsTop")				=0
		end if
		If IsHot=1 then
			rs("IsHot")				=1
		else
			rs("IsHot")				=0
		end if
		If IsFlash=1 then
			rs("IsFlash")			=1
		else
			rs("IsFlash")			=0
		end if
		If ArtDescription<>"" then
			rs("ArtDescription")		=ArtDescription
		Else
			rs("ArtDescription")		=Left(LoseHtml(Content),100)
		End if
		rs("TitleFontColor")	=TitleFontColor
		rs("Images")			=Images
		rs("yn")				=0
		rs("PageNum")			=PageNum
		rs("DateAndTime")		=DateAndTime
		rs("Vote")				=Vote
		rs("ReadPower")			=ReadPower

		If sSaveFileSelect="True" Then
      		Rs("Content")=ReplaceRemoteUrl(Content,SitePath&SiteUp&"/" & Year(Now) & right("0" & Month(Now), 2) & "/",sFileExt)
     	Else
      		Rs("Content")=Content
		End If
		
		rs.update
		Call Alert("�޸ĳɹ���","Admin_Article.asp")
	else
		Call Alert("�޸Ĵ���",-2)
	end if
	rs.close
	set rs=nothing
end sub

sub Admin_ShowClass_Option()
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from "&tbname&"_Class Where TopID = 0 and link=0 order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">��ѡ�����</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">������ӷ���</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option")
		 If int(session("YaoClassID"))=Rsp("ID") then
				Response.Write(" selected" ) 
		 End if
		 If Yao_MyID(rsp("ID"))<>"0" then
		 		Response.Write(" value=""-1"" style='background:#f7f7f7;color:#ccc;'")
		 Else
		 	If yaoadmintype<>1 then
				If instr(","&yaoWritePower&",",","& rsp("ID") &",") <> 0 or instr(","&yaoManagePower&",",","& rsp("ID") &",") <> 0 then
					Response.Write(" value=" & """" & Rsp("ID") & """" & " style='color:#0000ff;'")
				Else
					Response.Write(" value=""-1"" style='background:#f7f7f7;color:#ccc;'")
				End if
			else
				Response.Write(" value=" & """" & Rsp("ID") & """" & " style='color:#0000ff;'")
			End if
		 End if
         Response.Write(">|-" & Rsp("ClassName") & "</option>") & VbCrLf
		 
		    Sqlpp ="select * from "&tbname&"_Class Where TopID="&Rsp("ID")&" and link=0 order by num"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof 
				Response.Write("<option")
				If int(session("YaoClassID"))=Rspp("ID") then
				Response.Write(" selected" ) 
				End if
				If yaoadmintype<>1 then
					If instr(","&yaoWritePower&",",","& rspp("ID") &",") <> 0 or instr(","&yaoManagePower&",",","& rspp("ID") &",") <> 0 then
						Response.Write(" value=" & """" & Rspp("ID") & """" & " style='color:#0000ff;'")
					Else
						Response.Write(" value=""-1"" style='background:#f7f7f7;color:#ccc;'")
					End if
				else
					Response.Write(" value=" & """" & Rspp("ID") & """" & " style='color:#0000ff;'")
				End if
         		Response.Write(">��|-" & Rspp("ClassName") & "")
				Response.Write("</option>" ) 
			Rspp.Movenext   
      		Loop
			rspp.close
			set rspp=nothing
      Rsp.Movenext   
      Loop   
   End if
   rsp.close
   set rsp=nothing
end sub 

Sub delAll
ID=Trim(Request("ID"))
ytype=Request("ytype")
page=request("page")
If ID="" Then
	  Call Alert("��ѡ������!",-1)
ElseIf Request("Del")="����δ��" Then
   set rs=conn.execute("Update "&tbname&"_Article set yn = 1 where ID In(" & ID & ")")
ElseIf Request("Del")="����ʱ��" Then
   set rs=conn.execute("Update "&tbname&"_Article set DateAndTime = "&SqlNowString&" where ID In(" & ID & ")")
ElseIf Request("Del")="�������" Then
   set rs=conn.execute("Update "&tbname&"_Article set yn = 0 where ID In(" & ID & ")")
ElseIf Request("Del")="�Ƽ�" Then
   set rs=conn.execute("Update "&tbname&"_Article set IsHot = 1 where ID In(" & ID & ")")
ElseIf Request("Del")="����Ƽ�" Then
   set rs=conn.execute("Update "&tbname&"_Article set IsHot = 0 where ID In(" & ID & ")")
ElseIf Request("Del")="�̶�" Then
   set rs=conn.execute("Update "&tbname&"_Article set IsTop = 1 where ID In(" & ID & ")")
ElseIf Request("Del")="����̶�" Then
   set rs=conn.execute("Update "&tbname&"_Article set IsTop = 0 where ID In(" & ID & ")")
ElseIf Request("Del")="�ƶ�" Then
		If ytype="" or ytype="-1" then
			Call Alert("��ѡ����ȷ�����",-1)
		End if
   set rs=conn.execute("Update "&tbname&"_Article set ClassID = "&ytype&" where ID In(" & ID & ")")
ElseIf Request("Del")="ɾ��" Then
	'set rs=conn.execute("delete from "&tbname&"_Article where ID In(" & ID & ")")
			for i=1 to request("ID").count
				if request("ID").count=1 then
				ArticleID=request("ID")
				else
				ArticleID=replace(request("id")(i),"'","")
				end if
				'���û���ֵ����
				Call EditUserMn(ArticleID,money3,0)
				'ɾ������
				Conn.Execute("Delete From ["&tbname&"_Article] where ID = "&ArticleID&"")
				'ɾ�������������
				Conn.Execute("Delete From ["&tbname&"_Pl] where ArticleID = "&ArticleID&"")
			next
End If
Dim LaoyOK
LaoyOK=request.servervariables("http_referer")
If LaoyOK="" then LaoyOK=-1
Call Alert("�����ɹ�",LaoyOK)
End Sub
%>
<!--#include file="Admin_copy.asp"-->
<script language="javascript" src="../js/jscolor/jscolor.js"></script>
</body>
</html>