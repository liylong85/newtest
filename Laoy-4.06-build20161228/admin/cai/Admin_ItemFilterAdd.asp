<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<%

Dim Action,FoundErr,ErrMsg
Action=Trim(Request("Action"))
If Action="SaveAdd" Then
   Call Save()
Else
   Call Main()
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
<SCRIPT language=javascript>
function showset(thisform)
{
		if(thisform.FilterType.selectedIndex==1)
		{
        	FilterType1.style.display = "none";
			FilterType2.style.display = "";
		}
		else
		{
        	FilterType1.style.display = "";
			FilterType2.style.display = "none";
	    }

}
</script>
</head>
<body>
<form method="post" action="Admin_ItemFilterAdd.asp" name="form1">                                
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable" >                                
    <tr>                                 
      <td height="22" colspan="2" class="admintitle">�����¹���</td>                                
    </tr>                                
    <tr>                                 
      <td width="150" class="b1_1"><strong> �������ƣ�</strong></td>                                
      <td class="b1_1"><input name="FilterName" type="text" id="FilterName" size="25" maxlength="30">                                 
      &nbsp;</td>                                
    </tr>                                
    <tr>                                 
      <td width="150" class="b1_1"><strong> ������Ŀ��</strong></td>                                
      <td class="b1_1"><%Call Admin_ShowItem_Option(0)%>      </td>                                
    </tr>                                
    <tr>                                 
      <td width="150" class="b1_1"><strong>���˶���</strong></td>                               
      <td class="b1_1">                               
         <select name="FilterObject" id="FilterObject">                              
            <option value="1" selected>�������</option>                              
            <option value="2">���Ĺ���</option>                              
         </select>      </td>                               
    </tr>                               
    <tr>                                
      <td width="150" class="b1_1"><strong> �������ͣ�</strong></td>                               
      <td class="b1_1">                               
         <select name="FilterType" id="FilterType" onchange=showset(this.form)>                              
            <option value="1" selected >���滻</option>                              
            <option value="2">�߼�����</option>                              
         </select>      </td>                                
    </tr>    
    <tr>                                
      <td width="150" class="b1_1"><strong> ʹ��״̬��</strong></td>                               
      <td class="b1_1">                               
         <select name="Flag" id="Flag">                              
            <option value="yes" selected >����</option>                              
            <option value="no">����</option>                              
         </select>      </td>                                
    </tr>                              
    <tr>                                
      <td width="150" class="b1_1"><strong> ʹ�÷�Χ��</strong></td>                               
      <td class="b1_1">                               
         <select name="PublicTf" id="PublicTf">                              
            <option value="yes">����</option>                              
            <option value="no" selected>˽��</option>                              
         </select>      </td>                                
    </tr>
    <tr id="FilterType1">
      <td class="b1_1"><strong>���ݣ�</strong></td>
      <td class="b1_1"><textarea name="FilterContent" cols="49" rows="5"></textarea></td>
    </tr>
    <tr id="FilterType2" style="display:none;">
      <td colspan="2" class="b1_1"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">    <tr>                               
      <td width="150"><strong> ��ʼ��ǣ�</strong></td>                              
      <td><textarea name="FisString" cols="49" rows="5"></textarea>                               
        &nbsp;</td>                                
    </tr>                                
    <tr>                                 
      <td width="150"><strong> ������ǣ�</strong></td>                               
      <td><textarea name="FioString" cols="49" rows="5"></textarea>                                
        &nbsp;</td>                                
    </tr>                         
</table> </td>
    </tr>
    <tr>
      <td class="b1_1"><strong>�滻��</strong></td>
      <td class="b1_1"><textarea name="FilterRep" cols="49" rows="5"></textarea></td>
    </tr>
    <tr>
      <td class="b1_1">&nbsp;</td>
      <td class="b1_1"><input name="Action" type="hidden" id="Action" value="SaveAdd">
        <input name="Cancel" type="button" id="Cancel" value="ȡ&nbsp;&nbsp;��" onClick="window.location.href='Admin_ItemFilters.asp'">
&nbsp;&nbsp;&nbsp;&nbsp;
<input  type="submit" name="Submit" value="ȷ&nbsp;&nbsp;��"></td>
    </tr>                              
</table>                                                
</form>                               
<!--#include file="../Admin_Copy.asp"-->                                           
</body>                                         
</html>                                
<%End Sub%>                                
                                
<%                                
Sub Save()                                
                                
Dim SqlItem,RsItem                                
Dim FilterName,ItemID,FilterObject,FilterType,FilterContent,FisString,FioString,FilterRep,Flag,PublicTf 
FilterName=Trim(Request.Form("FilterName"))                                
ItemID=Trim(Request.Form("ItemID"))     
FilterObject=Request.Form("FilterObject")                              
FilterType=Request.Form("FilterType")                                
FilterContent=Request.Form("FilterContent")                                
FisString=Request.Form("FisString")                                
FioString=Request.Form("FioString")                                
FilterRep=Request.Form("FilterRep")                                
Flag=Request.Form("Flag")                              
PublicTf=Request.Form("PublicTf")                              

If FilterName="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>�������Ʋ���Ϊ��</li>"                                 
End If                                
If ItemID="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>��ѡ�����������Ŀ</li>"                                 
Else                                
   ItemID=Clng(ItemID)  
   If ItemID=0 Then  
      FoundErr=True                                
      ErrMsg=ErrMsg & "<br><li>��ѡ�����������Ŀ</li>"   
   End If                                
End If       
If FilterObject="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>��ѡ����˶���</li>"                                 
Else                                
   FilterObject=Clng(FilterObject)   
End If                                
                            
If FilterType="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>��ѡ���������</li>"                                 
Else                                
   FilterType=Clng(FilterType)                                
   If FilterType=1 Then     
      If FilterContent="" Then                              
         FoundErr=True                                
         ErrMsg=ErrMsg & "<br><li>���˵����ݲ���Ϊ��</li>"   
      End If   
   ElseIf FilterType=2 Then   
      If FisString="" or FioString="" Then   
         FoundErr=True                                
         ErrMsg=ErrMsg & "<br><li>��ʼ/������ǲ���Ϊ��</li>"   
      End If   
   Else   
      FoundErr=True                                
      ErrMsg=ErrMsg & "<br><li>�������������Ч���ӽ���</li>"   
   End If   
End If
If Flag="yes" Then
   Flag=True
Else
   Flag=False
End If     
If PublicTf="yes" Then
   PublicTf=True
Else
   PublicTf=False
End If                              
                                
If FoundErr<>True Then                                
   SqlItem ="select top 1 *  from Filters"                                
   Set RsItem=server.CreateObject("adodb.recordset")                                
   RsItem.open SqlItem,ConnItem,1,3                                
   RsItem.AddNew                                
   RsItem("FilterName")=FilterName                                
   RsItem("ItemID")=ItemID   
   RsItem("FilterObject")=FilterObject                                
   RsItem("FilterType")=FilterType                                
   If FilterType=1 Then                                
      RsItem("FilterContent")=FilterContent                                
   ElseIf FilterType=2 Then                                
      RsItem("FisString")=FisString                                
      RsItem("FioString")=FioString                                
   End If                                                
   RsItem("FilterRep")=FilterRep
   RsItem("Flag")=Flag 
   RsItem("PublicTf")=PublicTf                                   
   RsItem.Update                                
   RsItem.Close                                
   Set RsItem=Nothing                                
   Response.Redirect "Admin_ItemFilters.asp"                                
Else                                
   Call WriteErrMsg(ErrMsg)                                
End If                                
                                
End Sub                                
%>