<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
Dim Rs,Sql,FoundErr,ErrMsg
Dim SqlItem,RsItem
Dim ItemID,ItemName,WebName,WebUrl,ClassID,ChannelDir,SpecialID,ItemDemo,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse
Dim tClass,tSpecial

ItemName=Trim(Request.Form("ItemName"))
WebName=Trim(Request.Form("WebName"))
WebUrl=Trim(Request.Form("WebUrl"))
ClassID=Trim(Request.Form("ClassID"))
SpecialID=Trim(Request.Form("SpecialID"))
ItemDemo=Trim(Request.Form("ItemDemo"))
LoginType=Request.Form("LoginType")
LoginUrl=Trim(Request.Form("LoginUrl"))
LoginPostUrl=Trim(Request.Form("LoginPostUrl"))
LoginUser=Trim(Request.Form("LoginUser"))
LoginPass=Trim(Request.Form("LoginPass"))
LoginFalse=Trim(Request.Form("LoginFalse"))
ChannelDir=Trim(Request.Form("ChannelDir"))
If ItemName="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>��Ŀ���Ʋ���Ϊ��</li>"
End If
If WebName="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>��վ���Ʋ���Ϊ��</li>"
End If
If WebUrl="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>��վ��ַ����Ϊ��</li>" 
End If
If ClassID=""  Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>δָ����Ŀ</li>"
Else
   ClassID=CLng(ClassID)
   set rs=conn.execute("select * from "&tbname&"_Class Where ID="  & ClassID)
   If rs.bof and rs.eof then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ</li>"
    End If
	Set rs=Nothing
End if

If LoginType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>��ѡ���¼����</li>"
Else
   LoginType=Clng(LoginType)
   If LoginType=1 Then
         If LoginUrl="" or LoginPostUrl="" or LoginUser="" or  LoginPass="" or  LoginFalse="" then
         FoundErr=True
         ErrMsg=ErrMsg& "<br><li>�뽫��¼������д����</li>"
      End If
   End If
End If

If FoundErr<>True Then
   SqlItem="Select top 1 ItemID,ItemName,WebName,WebUrl,ClassID,ChannelDir,ClassID,SpecialID,ItemDemo,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse from Item"
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,3
   RsItem.AddNew
   RsItem("ItemName")=ItemName
   RsItem("WebName")=WebName
   RsItem("WebUrl")=WebUrl
   RsItem("ClassID")=ClassID
   RsItem("ChannelDir")=ChannelDir
   If ItemDemo<>"" then
      RsItem("ItemDemo")=ItemDemo
   End if
   RsItem("LoginType")=LoginType
   If LoginType=1 Then
      RsItem("LoginUrl")=LoginUrl
      RsItem("LoginPostUrl")=LoginPostUrl
      RsItem("LoginUser")=LoginUser
      RsItem("LoginPass")=LoginPass
      RsItem("LoginFalse")=LoginFalse
   End If
   ItemID=RsItem("ItemID")
   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing
End If

If FoundErr=True Then
   call WriteErrMsg(ErrMsg)
Else
   Call Main
End If
'�ر����ݿ�����
Call CloseConn()
Call CloseConnItem()
%>
<%Sub Main%>
<html>
<head>
<title>���²ɼ�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../images/Admin_css.css">
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="30" class="b1_1"><a href="Admin_ItemAddNew.asp">������Ŀ</a> >> <a href="Admin_ItemModify.asp">��������</a> >> <font color=red>�б�����</font> >> �������� >> �������� >> �������� >> �������� >> ���</td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable" >
<form method="post" action="Admin_ItemAddNew3.asp" name="form1">
    <tr> 
      <td height="22" colspan="2" class="admintitle">��������Ŀ--�б�����</td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1"><strong>�б�����ҳ�棺</strong></td>
      <td width="75%" class="b1_1">
		<input name="ListStr" type="text" size="58" maxlength="200">      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>�б���ʼ��ǣ�</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="LsString" cols="49" rows="7"></textarea><br>      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>�б�������ǣ�</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="LoString" cols="49" rows="7"></textarea><br>      </td>
    </tr>
    <tr>
      <td width="20%" class="b1_1"><strong> �б�������ҳ��</strong></td>
      <td width="75%" class="b1_1">
		<input name="ListPaingType" type="radio" class="noborder" onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='none';ListPaing3.style.display='none'" value="0" checked>
		��������&nbsp;
		<input name="ListPaingType" type="radio" class="noborder" onClick="ListPaing1.style.display='';ListPaing12.style.display='';ListPaing2.style.display='none';ListPaing3.style.display='none'" value="1">
		���ñ�ǩ&nbsp;
		<input name="ListPaingType" type="radio" class="noborder" onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='';ListPaing3.style.display='none'" value="2">
		��������&nbsp;
		<input name="ListPaingType" type="radio" class="noborder" onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='none';ListPaing3.style.display=''" value="3">
		�ֶ�����      </td>
    </tr>
	<tr id="ListPaing1" style="display:none">
      <td width="20%" class="b1_1"><strong><font color=blue>��ҳ��ʼ��ǣ�</font></strong>
        <p>��</p><p>��</p>
      <strong><font color=blue>��ҳ������ǣ�</font></strong>      </td>
      <td width="75%" class="b1_1">
		<textarea name="LPsString" cols="49" rows="7"></textarea><br>
		<textarea name="LPoString" cols="49" rows="7"></textarea>      </td>
    </tr>
	<tr id="ListPaing12" style="display:none">
      <td width="20%" class="b1_1"><strong><font color=blue>������ҳ�ض���</font></strong>      </td>
      <td width="75%" class="b1_1">
	<input name="ListPaingStr1" type="text" size="58" maxlength="200" value=""><br>
        ��ʽ��http://www.laoy.net/list.asp?page={$ID}      </td>
    </tr>
    <tr id="ListPaing2" style="display:none"> 
      <td width="20%" class="b1_1"><strong><font color=blue>�������ɣ�</font></strong></td>
      <td width="75%" class="b1_1">
        ԭ�ַ�����<br>
		<input name="ListPaingStr2" type="text" size="58" maxlength="200" value=""><br>
                ��ʽ��http://www.laoy.net/list.asp?page={$ID}<br>
                <br>
	    ���ɷ�Χ��<br>
		<input name="ListPaingID1" type="text" size="8" maxlength="200"><span lang="en-us"> To </span><input name="ListPaingID2" type="text" size="8" maxlength="200"><br>
                ��ʽ��ֻ�������֣���������߽���      </td>
    </tr>
    <tr id="ListPaing3" style="display:none"> 
      <td width="20%" class="b1_1"><strong><font color=blue>�ֶ����ӣ�</font></strong>      </td>
      <td width="75%" class="b1_1">
      <textarea name="ListPaingStr3" cols="49" rows="7"></textarea><br>
      ��ʽ������һ����ַ�󰴻س�����������һ����      </td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="b1_1">
        <input name="ItemID" type="hidden" value="<%=ItemID%>">
      <input  type="submit" name="Submit" value="��&nbsp;һ&nbsp;��"></td>
    </tr>
</form>
</table>
<!--#include file="../Admin_Copy.asp"-->        
</body>         
</html>
<%End Sub%>