<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
Dim ItemID
Dim RsItem,SqlItem,FoundErr,ErrMsg
Dim ListStr,LsString,LoString
Dim  ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr,ChannelDir
Dim  ListUrl,ListCode,NewsArrayCode,NewsArray,UrlTest,NewsCode
dim Testi
ItemID=Trim(Request("ItemID"))
HsString=Request.Form("HsString")
HoString=Request.Form("HoString")
HttpUrlType=Trim(Request.Form("HttpUrlType"))
HttpUrlStr=Trim(Request.Form("HttpUrlStr"))


If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>�������������Ч���ӽ���</li>"
Else
   ItemID=Clng(ItemID)
End If
If HsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>���ӿ�ʼ��ǲ���Ϊ��</li>"
End If
If HoString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>���ӽ�����ǲ���Ϊ��</li>" 
End If
If HttpUrlType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>��ѡ�����Ӵ�������</li>"
Else
   HttpUrlType=Clng(HttpUrlType)
   If HttpUrlType=1 Then
      If HttpUrlStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>���������ַ����ò���Ϊ��</li>"
      Else
            If  Len(HttpUrlStr)<15  Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>���������ַ����ò���ȷ(15���ַ�����)</li>"
            End  If      
      End If
   End If
End If


If FoundErr<>True Then
   SqlItem="Select * from Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,2,3

   RsItem("HsString")=HsString
   RsItem("HoString")=HoString
   RsItem("HttpUrlType")=HttpUrlType
   If HttpUrlType=1 Then
      RsItem("HttpUrlStr")=HttpUrlStr
   End If
   ListStr=RsItem("ListStr")
   LsString=RsItem("LsString")
   LoString=RsItem("LoString")
   ListPaingType=RsItem("ListPaingType")
   LPsString=RsItem("LPsString")
   ListPaingStr1=RsItem("ListPaingStr1")
   ListPaingStr2=RsItem("ListPaingStr2")
   ListPaingID1=RsItem("ListPaingID1")
   ListPaingID2=RsItem("ListPaingID2")
   ListPaingStr3=RsItem("ListPaingStr3")
   ChannelDir=RsItem("ChannelDir")

   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing
   
   Select  Case  ListPaingType
   Case  0,1
         ListUrl=ListStr
   Case  2
         ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1))
   Case  3
         If  Instr(ListPaingStr3,"|")>0  Then
         ListUrl=Left(ListPaingStr3,Instr(ListPaingStr3,"|")-1)
      Else
         ListUrl=ListPaingStr3
      End  If
   End  Select

   ListCode=GetHttpPage(ListUrl,ChannelDir)
   If ListCode<>"$False$" Then
      ListCode=GetBody(ListCode,LsString,LoString,False,False)
      If ListCode<>"$False$" Then 
         NewsArrayCode=GetArray(ListCode,HsString,HoString,False,False)
         If NewsArrayCode<>"$False$" Then
               NewsArray=Split(NewsArrayCode,"$Array$")
               For Testi=0 To Ubound(NewsArray)
                  If HttpUrlType=1 Then
                     NewsArray(Testi)=Replace(HttpUrlStr,"{$ID}",NewsArray(Testi))
                  Else
                     NewsArray(Testi)=DefiniteUrl(NewsArray(Testi),ListUrl)
                  End If
               Next
               UrlTest=NewsArray(0)
               NewsCode=GetHttpPage(UrlTest,ChannelDir)
         Else
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ�����б�ʱ������</li>"
         End If   
      Else
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>�ڽ�ȡ�б�ʱ��������</li>"
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & ListUrl & "��ҳԴ��ʱ��������</li>"
   End If
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End If 
'�ر����ݿ�����
Call CloseConn()
Call CloseConnItem()
%>
<%Sub Main()%>
<html>
<head>
<title>�ɼ�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../images/Admin_css.css">
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="30" class="b1_1"><a href="Admin_ItemAddNew.asp">������Ŀ</a> >> <a href="Admin_ItemModify.asp?ItemID=<%=ItemID%>">��������</a> >> <a href="Admin_ItemModify2.asp?ItemID=<%=ItemID%>">�б�����</a> >> <a href="Admin_ItemModify3.asp?ItemID=<%=ItemID%>">��������</a> >> <font color=red>��������</font> >> �������� >> �������� >> ���</td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="admintable" >        
    <tr> 
      <td height="22" colspan="2" class="admintitle">��������Ŀ--�б��������Ӳ���</td>
    </tr>
    <tr>
      <td height="22" colspan="2" class="b1_1">�����Ƿ��������õ������¾������ӵ�ַ����鿴�Ƿ���ȷ��<br>
        <%
For Testi=0 To Ubound(NewsArray)
   Response.Write "<a href='" & NewsArray(Testi) & "' target=_blank>" & NewsArray(Testi) & "</a><br>"
Next
%>
        <br>
��һ������ȡ��һ�����½��в��ԣ�����д���±��ʱ������Ҫʹ�õ�һ�����¡�</td>
    </tr>
</table>
<form method="post" action="Admin_ItemAddNew5.asp" name="form1">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable" >
    <tr> 
      <td height="22" colspan="2" class="admintitle">��������Ŀ--��������      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>���⿪ʼ��ǣ�</strong>
        <p>��</p><p>��</p>
      <strong>���������ǣ�</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="TsString" cols="49" rows="7"></textarea><br>
      <textarea name="ToString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>���Ŀ�ʼ��ǣ�</strong>
        <p>��</p><p>��</p>
      <strong>���Ľ�����ǣ�</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="CsString" cols="49" rows="7"></textarea><br>
      <textarea name="CoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><b>ʱ<span lang="en-us">&nbsp;</span>��<span lang="en-us">&nbsp;</span>��<span lang="en-us">&nbsp;</span>�ã�</b></td>
      <td width="75%" class="b1_1">
      	<input name="DateType" type="radio" class="noborder" onClick="Date1.style.display='none'" value="0" checked>
      	��������&nbsp;
		<input name="DateType" type="radio" class="noborder" onClick="Date1.style.display=''" value="1">
		���ñ�ǩ&nbsp;
    </tr>
    <tr id="Date1" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>ʱ�俪ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>ʱ�������ǣ�</font></strong></td>
      <td width="75%" class="b1_1">
      <textarea name="DsString" cols="49" rows="7"></textarea><br>
      <textarea name="DoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><b>�� ��<span lang="en-us">&nbsp;</span>��<span lang="en-us">&nbsp;</span>�ã�</b></td>
      <td width="75%" class="b1_1">
      	<input name="AuthorType" type="radio" class="noborder" onClick="Author1.style.display='none';Author2.style.display='none'" value="0" checked>
      	��������&nbsp;
		<input name="AuthorType" type="radio" class="noborder" onClick="Author1.style.display='';Author2.style.display='none'" value="1">
		���ñ�ǩ&nbsp;
	  <input name="AuthorType" type="radio" class="noborder" onClick="Author1.style.display='none';Author2.style.display=''" value="2">
	  ָ������</td>
    </tr>
    <tr id="Author1" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>���߿�ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>���߽�����ǣ�</font></strong></td>
      <td width="75%" class="b1_1">
      <textarea name="AsString" cols="49" rows="7"></textarea><br>
      <textarea name="AoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr id="Author2" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>��ָ�����ߣ�</font></strong></td>
      <td width="75%" class="b1_1">
      <input name="AuthorStr" type="text" id="AuthorStr" value="">
      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><b>�� Դ&nbsp;�� �ã�</b></td>
      <td width="75%" class="b1_1">
      	<input name="CopyFromType" type="radio" class="noborder" onClick="CopyFrom1.style.display='none';CopyFrom2.style.display='none'" value="0" checked>
      	��������&nbsp;
		<input name="CopyFromType" type="radio" class="noborder" onClick="CopyFrom1.style.display='';CopyFrom2.style.display='none'" value="1">
		���ñ�ǩ&nbsp;
	  <input name="CopyFromType" type="radio" class="noborder" onClick="CopyFrom1.style.display='none';CopyFrom2.style.display=''" value="2">
	  ָ����Դ</td>
    </tr>
    <tr id="CopyFrom1" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>��Դ��ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>��Դ������ǣ�</font></strong></td>
      <td width="75%" class="b1_1">
      <textarea name="FsString" cols="49" rows="7"></textarea><br>
      <textarea name="FoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr id="CopyFrom2" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>��ָ����Դ��</font></strong></td>
      <td width="75%" class="b1_1">
      <input name="CopyFromStr" type="text" id="CopyFromStr" value="">
      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><b>�ؼ��ִ����ã�</b></td>
      <td width="75%" class="b1_1">
      	<input name="KeyType" type="radio" class="noborder" onClick="Key1.style.display='none';Key2.style.display='none'" value="3" checked>
      	���ݴʿ�����[�Բɼ��ٶ���Ӱ��]&nbsp;
      	<input name="KeyType" type="radio" class="noborder" onClick="Key1.style.display='none';Key2.style.display='none'" value="0">
      	��������&nbsp;
		<input name="KeyType" type="radio" class="noborder" onClick="Key1.style.display='';Key2.style.display='none'" value="1">
		��ǩ����&nbsp;
	  <input name="KeyType" type="radio" class="noborder" onClick="Key1.style.display='none';Key2.style.display=''" value="2">
	  �Զ���ؼ���</td>
    </tr>
    <tr id="Key1" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>�ؼ��ʿ�ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>�ؼ��ʽ�����ǣ�</font></strong></td>
      <td width="75%" class="b1_1">
      <textarea name="KsString" cols="49" rows="7"></textarea><br>
      <textarea name="KoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr id="Key2" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>��ָ���ؼ���</font></strong></td>
      <td width="75%" class="b1_1">
      <input name="KeyStr" type="text" id="KeyStr" value="">
      <span class="note">��ʽ���ؼ���֮����<font color=red>|</font>�ָ����磺����|IT</span>
      </td>
    </tr>
    <tr>
      <td width="20%" class="b1_1"><strong>���ķ�ҳ���ã�</strong></td>
      <td width="75%" class="b1_1">
		<input name="NewsPaingType" type="radio" class="noborder" onClick="NewsPaing1.style.display='none';NewsPaing12.style.display='none';NewsPaing13.style.display='none';NewsPaing2.style.display='none'" value="0" checked>
		��������&nbsp;
		<input name="NewsPaingType" type="radio" class="noborder" onClick="NewsPaing1.style.display='';NewsPaing12.style.display='';NewsPaing13.style.display='';NewsPaing2.style.display='none'" value="1">
		���ñ�ǩ&nbsp;
		<input name="NewsPaingType" type="radio" class="noborder" onClick="NewsPaing1.style.display='none';NewsPaing12.style.display='none';NewsPaing13.style.display='none';NewsPaing2.style.display=''" value="2">
		��������
      </td>
    </tr>
	<tr id="NewsPaing1" style="display:none">
      <td width="20%" class="b1_1"><strong><font color=blue>��ҳ��ʼ��ǣ�</font></strong>
        <p>��</p><p>��</p>
      <strong><font color=blue>��ҳ������ǣ�</font></strong></td>
      <td width="75%" class="b1_1">
		<textarea name="NPsString" cols="49" rows="7"></textarea><br>
	  <textarea name="NPoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr id="NewsPaing12" style="display:none"> 
      <td width="20%" class="b1_1"><b><font color="#0000FF">��ҳ�������ӣ�</font></b></td>
      <td width="75%" class="b1_1">
		<input name="NewsPaingStr" type="text" size="58"><br>
		��ʽ��http://www.laoy.net/Article_Show.asp?ID=1&Page={$ID}
    </tr>
    <tr id="NewsPaing13" style="display:'none'"> 
      <td width="20%" class="b1_1"><b><font color="#0000FF">��ҳ�����ַ���</font></b></td>
      <td width="75%" class="b1_1">
	  <input name="NewsPaingHtml" type="text" size="58" value=""></td>
    </tr>
    <tr id="NewsPaing2" style="display:none"> 
      <td width="20%" class="b1_1"><strong><font color=blue>��&nbsp; ��&nbsp; ��&nbsp; �ã�</font></strong></td>
      <td width="75%" class="b1_1">
		<input name="NewsPaingStr2" type="text" value="Ԥ������" size="51">
      </td>
    </tr>

    <tr> 
      <td colspan="2" align="center" class="b1_1"><br>
          <input name="ItemID" type="hidden" value="<%=ItemID%>"> 
        <input  type="button" name="button1" value="��&nbsp;һ&nbsp;��" onClick="window.location.href='javascript:history.go(-1)'" >
        &nbsp;&nbsp;&nbsp;&nbsp; 
      <input  type="submit" name="Submit" value="��&nbsp;һ&nbsp;��">
      <input name="ChannelDir" type="hidden" id="ChannelDir" value="<%=ChannelDir%>"></td>
        <input type="hidden" name="UrlTest" id="UrlTest" value="<%=UrlTest%>">
    </tr>
</table>
</form>
<!--#include file="../Admin_Copy.asp"-->        
</body>         
</html>
<%End Sub%>