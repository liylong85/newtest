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
   ErrMsg=ErrMsg & "<br><li>项目名称不能为空</li>"
End If
If WebName="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>网站名称不能为空</li>"
End If
If WebUrl="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>网站网址不能为空</li>" 
End If
If ClassID=""  Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>未指定栏目</li>"
Else
   ClassID=CLng(ClassID)
   set rs=conn.execute("select * from "&tbname&"_Class Where ID="  & ClassID)
   If rs.bof and rs.eof then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>找不到指定的栏目</li>"
    End If
	Set rs=Nothing
End if

If LoginType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请选择登录类型</li>"
Else
   LoginType=Clng(LoginType)
   If LoginType=1 Then
         If LoginUrl="" or LoginPostUrl="" or LoginUser="" or  LoginPass="" or  LoginFalse="" then
         FoundErr=True
         ErrMsg=ErrMsg& "<br><li>请将登录参数填写完整</li>"
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
'关闭数据库链接
Call CloseConn()
Call CloseConnItem()
%>
<%Sub Main%>
<html>
<head>
<title>文章采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../images/Admin_css.css">
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="30" class="b1_1"><a href="Admin_ItemAddNew.asp">添加项目</a> >> <a href="Admin_ItemModify.asp">基本设置</a> >> <font color=red>列表设置</font> >> 链接设置 >> 正文设置 >> 采样测试 >> 属性设置 >> 完成</td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable" >
<form method="post" action="Admin_ItemAddNew3.asp" name="form1">
    <tr> 
      <td height="22" colspan="2" class="admintitle">添加新项目--列表设置</td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1"><strong>列表索引页面：</strong></td>
      <td width="75%" class="b1_1">
		<input name="ListStr" type="text" size="58" maxlength="200">      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>列表开始标记：</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="LsString" cols="49" rows="7"></textarea><br>      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>列表结束标记：</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="LoString" cols="49" rows="7"></textarea><br>      </td>
    </tr>
    <tr>
      <td width="20%" class="b1_1"><strong> 列表索引分页：</strong></td>
      <td width="75%" class="b1_1">
		<input name="ListPaingType" type="radio" class="noborder" onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='none';ListPaing3.style.display='none'" value="0" checked>
		不作设置&nbsp;
		<input name="ListPaingType" type="radio" class="noborder" onClick="ListPaing1.style.display='';ListPaing12.style.display='';ListPaing2.style.display='none';ListPaing3.style.display='none'" value="1">
		设置标签&nbsp;
		<input name="ListPaingType" type="radio" class="noborder" onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='';ListPaing3.style.display='none'" value="2">
		批量生成&nbsp;
		<input name="ListPaingType" type="radio" class="noborder" onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='none';ListPaing3.style.display=''" value="3">
		手动添加      </td>
    </tr>
	<tr id="ListPaing1" style="display:none">
      <td width="20%" class="b1_1"><strong><font color=blue>下页开始标记：</font></strong>
        <p>　</p><p>　</p>
      <strong><font color=blue>下页结束标记：</font></strong>      </td>
      <td width="75%" class="b1_1">
		<textarea name="LPsString" cols="49" rows="7"></textarea><br>
		<textarea name="LPoString" cols="49" rows="7"></textarea>      </td>
    </tr>
	<tr id="ListPaing12" style="display:none">
      <td width="20%" class="b1_1"><strong><font color=blue>索引分页重定向：</font></strong>      </td>
      <td width="75%" class="b1_1">
	<input name="ListPaingStr1" type="text" size="58" maxlength="200" value=""><br>
        格式：http://www.laoy.net/list.asp?page={$ID}      </td>
    </tr>
    <tr id="ListPaing2" style="display:none"> 
      <td width="20%" class="b1_1"><strong><font color=blue>批量生成：</font></strong></td>
      <td width="75%" class="b1_1">
        原字符串：<br>
		<input name="ListPaingStr2" type="text" size="58" maxlength="200" value=""><br>
                格式：http://www.laoy.net/list.asp?page={$ID}<br>
                <br>
	    生成范围：<br>
		<input name="ListPaingID1" type="text" size="8" maxlength="200"><span lang="en-us"> To </span><input name="ListPaingID2" type="text" size="8" maxlength="200"><br>
                格式：只能是数字，可升序或者降序。      </td>
    </tr>
    <tr id="ListPaing3" style="display:none"> 
      <td width="20%" class="b1_1"><strong><font color=blue>手动添加：</font></strong>      </td>
      <td width="75%" class="b1_1">
      <textarea name="ListPaingStr3" cols="49" rows="7"></textarea><br>
      格式：输入一个网址后按回车，再输入下一个。      </td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="b1_1">
        <input name="ItemID" type="hidden" value="<%=ItemID%>">
      <input  type="submit" name="Submit" value="下&nbsp;一&nbsp;步"></td>
    </tr>
</form>
</table>
<!--#include file="../Admin_Copy.asp"-->        
</body>         
</html>
<%End Sub%>