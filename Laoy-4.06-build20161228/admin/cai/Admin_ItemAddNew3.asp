<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
Dim SqlItem,RsItem,ItemID,FoundErr,ErrMsg
Dim ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,ChannelDir
Dim ListUrl,ListCode
Dim LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,LoginResult,LoginData
Dim  ListPaingNext
ItemID=Trim(Request.Form("ItemID"))
ListStr=Trim(Request.Form("ListStr"))
LsString=Request.Form("LsString")
LoString=Request.Form("LoString")
ChannelDir=Request.Form("ChannelDir")
ListPaingType=Request.Form("ListPaingType")
LPsString=Request.Form("LPsString")
LPoString=Request.Form("LPoString")
ListPaingStr1=Trim(Request.Form("ListPaingStr1"))
ListPaingStr2=Trim(Request.Form("ListPaingStr2"))
ListPaingID1=Trim(Request.Form("ListPaingID1"))
ListPaingID2=Trim(Request.Form("ListPaingID2"))
ListPaingStr3=Trim(Request.Form("ListPaingStr3"))
FoundErr=False

If ItemID=""  Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>�������������Ч���ӽ���</li>"
Else
   ItemID=Clng(ItemID)
End If
If LsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>�б���ʼ��ǲ���Ϊ��</li>"
End If
If LoString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>�б�������ǲ���Ϊ��</li>" 
End If
If ListPaingType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>��ѡ���б�������ҳ����</li>" 
Else
   ListPaingType=Clng(ListPaingType)
   Select Case ListPaingType
   Case 0,1
            If ListStr="" Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>�б�����ҳ����Ϊ��</li>"
            Else
               ListStr=Trim(ListStr)
            End If
      If  ListPaingType=1  Then
            If LPsString="" or LPoString="" Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>������ҳ��ʼ/������ǲ���Ϊ��</li>" 
            End If
            If ListPaingStr1<>"" and Len(ListPaingStr1)<15 Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>������ҳ�ض������ò���ȷ(����15���ַ�)</li>" 
            End If
      End  If
   Case 2
      If ListPaingStr2="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>���������ַ�����Ϊ��</li>"
      End If
      If isNumeric(ListPaingID1)=False or isNumeric(ListPaingID2)=False Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>�������ɵķ�Χֻ��������</li>"
      Else
         ListPaingID1=Clng(ListPaingID1)
         ListPaingID2=Clng(ListPaingID2)
         If ListPaingID1=0 And ListPaingID2=0 Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>�������ɷ�Χ���ò���ȷ</li>"
         End If
      End If
   Case 3
      If ListPaingStr3="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>�б�������ҳ����Ϊ�գ����ֶ�����</li>"
      Else
         ListPaingStr3=Replace(ListPaingStr3,CHR(13),"|") 
      End If
   Case Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>��ѡ���б�������ҳ����</li>" 
   End Select
End if

If FoundErr<>True Then
   SqlItem="Select * from Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,2,3

   RsItem("LsString")=LsString
   RsItem("LoString")=LoString
   RsItem("ListPaingType")=ListPaingType
   Select Case ListPaingType
   Case 0,1
         RsItem("ListStr")=ListStr
      If ListPaingType=1  Then
            RsItem("LPsString")=LPsString
            RsItem("LPoString")=LPoString
            If ListPaingStr1<>"" Then
               RsItem("ListPaingStr1")=ListPaingStr1
            End If
      End  If
      ListUrl=ListStr
   Case 2
      RsItem("ListPaingStr2")=ListPaingStr2
      RsItem("ListPaingID1")=ListPaingID1
      RsItem("ListPaingID2")=ListPaingID2
      ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1))
   Case 3
      RsItem("ListPaingStr3")=ListPaingStr3
      If  Instr(ListPaingStr3,"|")>0  Then
            ListUrl=Left(ListPaingStr3,Instr(ListPaingStr3,"|")-1)
      Else
            ListUrl=ListPaingStr3
      End  If
   End Select
   LoginType=RsItem("LoginType")
   LoginUrl=RsItem("LoginUrl")
   LoginPostUrl=RsItem("LoginPostUrl")
   LoginUser=RsItem("LoginUser")
   LoginPass=RsItem("LoginPass")
   LoginFalse=RsItem("LoginFalse")
   ChannelDir=RsItem("ChannelDir")
   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing

   If LoginType=1 then
      LoginData=UrlEncoding(LoginUser & "&" & LoginPass)
      LoginResult=PostHttpPage(LoginUrl,LoginPostUrl,LoginData)
      If Instr(LoginResult,LoginFalse)>0 Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>��վ��¼ʧ�ܣ������¼������</li>"
      End If
   End If
   
   If FoundErr<>True Then
      ListCode=GetHttpPage(ListUrl,ChannelDir)
      If ListCode<>"$False$" Then
         If ListPaingType=1  Then
            ListPaingNext=GetPaing(ListCode,LPsString,LPoString,False,False)
                  If ListPaingNext<>"$False$"  Then
                     If ListPaingStr1<>""  Then  
                        ListPaingNext=Replace(ListPaingStr1,"{$ID}",ListPaingNext)
               Else
                        ListPaingNext=DefiniteUrl(ListPaingNext,ListUrl)
               End  If
            End  If
         End If
         ListCode=GetBody(ListCode,LsString,Lostring,False,False)
         If ListCode="$False$" Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>�ڽ�ȡ:" & ListUrl & "�����б�ʱ��������</li>"
         End If
      Else
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & ListUrl & "��ҳԴ��ʱ��������</li>"
      End If
   End If
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
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
<title>�ɼ�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../images/Admin_css.css">
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="30" class="b1_1"><a href="Admin_ItemAddNew.asp">������Ŀ</a> >> <a href="Admin_ItemModify.asp?ItemID=<%=ItemID%>">��������</a> >> <a href="Admin_ItemModify2.asp?ItemID=<%=ItemID%>">�б�����</a> >> <font color=red>��������</font> >> �������� >> �������� >> �������� >> ���</td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
    <tr> 
      <td height="22" colspan="2" class="admintitle">��������Ŀ--�б���ȡ����</td>
    </tr>
    <tr> 
      <td height="22" colspan="2">
	  <textarea name="Content" id="Content" style="width:100%;height:300px;"><%=ListCode%></textarea>
      </td>
    </tr>
</table>
<%If ListPaingNext<>"" And ListPaingNext<>"$False$" Then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="admintable">
    <tr> 
      <td height="22" colspan="2" >
      <%Response.Write "<br>��һҳ�б���<a  href='" & ListPaingNext  &  "' target=_blank><font  color=red>"  &  ListPaingNext  &  "</font></a>"%>
      </td>
    </tr>
</table>
<%End If%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable" >
<form method="post" action="Admin_ItemAddNew4.asp" name="form1">
    <tr> 
      <td colspan="2" class="admintitle">��������Ŀ--��������</td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>���ӿ�ʼ��ǣ�</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="HsString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>���ӽ�����ǣ�</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="HoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr>
      <td width="20%" class="b1_1"><strong>���Ӵ������ͣ�</strong></td>
      <td width="75%" class="b1_1">
		<input name="HttpUrlType" type="radio" class="noborder" onClick="HttpUrl1.style.display='none'" value="0" checked>
		�Զ�����&nbsp;
		<input name="HttpUrlType" type="radio" class="noborder" onClick="HttpUrl1.style.display=''" value="1">
		���¶���
      </td>
    </tr>
	<tr id="HttpUrl1" style="display:none">
      <td width="20%" class="b1_1"><strong>���¶��������ַ���</strong></td>
      <td width="75%" class="b1_1">
	<input name="HttpUrlStr" type="text" size="49" maxlength="200" value=""><br>
        ��ʽ��http://www.laoy.nett/Article_Show.asp?ID={$ID}
      </td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="b1_1">
        <input name="ItemID" type="hidden" value="<%=ItemID%>">
        <input  type="button" name="button1" value="��&nbsp;һ&nbsp;��" onClick="window.location.href='javascript:history.go(-1)'" >
        &nbsp;&nbsp;&nbsp;&nbsp; 
      <input  type="submit" name="Submit" value="��&nbsp;һ&nbsp;��"></td>
    </tr>
</form>
</table>
<!--#include file="../Admin_Copy.asp"-->        
</body>         
</html>
<%End Sub%>