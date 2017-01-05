<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
Dim Rs,Sql,FoundErr,ErrMsg,Action
Dim SqlItem,RsItem
Dim ItemID,ItemName,WebName,WebUrl,ClassID,ChannelDir,SpecialID,ItemDemo,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse
Dim ListUrl,LsString,LoString,ListPaingType,LPsString,LPoString,ListStr,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3
Dim tClass,tSpecial
FoundErr=False

Action=Trim(Request("Action"))
ItemID=Trim(Request("ItemID"))

If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，项目ID不能为空！</li>"
Else
   ItemID=Clng(ItemID)
End If

If Action="SaveEdit" And FoundErr<>True Then
   Call SaveEdit()
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End If
'关闭数据库链接
Call CloseConn()
Call CloseConnItem()
%>
<%Sub Main
   SqlItem="Select * from Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If RsItem.Eof And RsItem.Bof Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>没有找到该项目!</li>"
   Else
      LoginType=RsItem("LoginType")
      LoginUrl=RsItem("LoginUrl")
      LoginPostUrl=RsItem("LoginPostUrl")
      LoginUser=RsItem("LoginUser")
      LoginPass=RsItem("LoginPass")
      LoginFalse=RsItem("LoginFalse")
      ListStr=RsItem("ListStr")
      LsString=RsItem("LsString")
      LoString=RsItem("LoString")
      ListPaingType=RsItem("ListPaingType")
      LPsString=RsItem("LPsString")
      LPoString=RsItem("LPoString")
      ListPaingStr1=RsItem("ListPaingStr1")
      ListPaingStr2=RsItem("ListPaingStr2")
      ListPaingID1=RsItem("ListPaingID1")
      ListPaingID2=RsItem("ListPaingID2")
      ListPaingStr3=RsItem("ListPaingStr3")
	  ChannelDir=RsItem("ChannelDir")
      If ListPaingStr3<>"" Then
         ListPaingStr3=Replace(ListPaingStr3,"|",CHR(13))
      End If
   End If
   RsItem.Close
   Set RsItem=Nothing
   If FoundErr=True Then
      Call WriteErrMsg(ErrMsg)

   Else
%>
<html>
<head>
<title>采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../images/Admin_css.css">
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="30" class="b1_1"><a href="Admin_ItemAddNew.asp">添加项目</a> >> <a href="Admin_ItemModify.asp?ItemID=<%=ItemID%>">基本设置</a> >> <font color=red>列表设置</font> >> 链接设置 >> 正文设置 >> 采样测试 >> 属性设置 >> 完成</td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable" >
<form method="post" action="Admin_ItemModify3.asp" name="form1">
    <tr> 
      <td height="22" colspan="2" class="admintitle">编辑项目--列表设置</td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1"><strong>列表索引页面：</strong></td>
      <td width="75%" class="b1_1">
		<input name="ListStr" type="text" size="58" maxlength="200" value="<%=ListStr%>">&nbsp;&nbsp;列表的第一页      </td>
    </tr>
    <tr>
      <td class="b1_1"><strong>页面编码类型：</strong></td>
      <td width="75%" class="b1_1"><select name="ChannelDir" id="ChannelDir">
        <option value="gb2312"<%If ChannelDir="gb2312" then Response.Write(" selected") end if%>>gb2312</option>
        <option value="utf-8"<%If ChannelDir="utf-8" then Response.Write(" selected") end if%>>utf-8</option>
      </select>
      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>列表开始标记：</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="LsString" cols="49" rows="7"><%=LsString%></textarea><br>      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>列表结束标记：</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="LoString" cols="49" rows="7"><%=LoString%></textarea><br>      </td>
    </tr>

    <tr>
      <td width="20%" class="b1_1"><strong> 列表索引分页：</strong></td>
      <td width="75%" class="b1_1">
		<input name="ListPaingType" type="radio" class="noborder" onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='none';ListPaing3.style.display='none'" value="0" <%If ListPaingType=0 Then Response.Write "checked"%>>
		不作设置&nbsp;
		<input name="ListPaingType" type="radio" class="noborder" onClick="ListPaing1.style.display='';ListPaing12.style.display='';ListPaing2.style.display='none';ListPaing3.style.display='none'" value="1" <%If ListPaingType=1 Then Response.Write "checked"%>>
		设置标签&nbsp;
		<input name="ListPaingType" type="radio" class="noborder" onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='';ListPaing3.style.display='none'" value="2" <%If ListPaingType=2 Then Response.Write "checked"%>>
		批量生成&nbsp;
		<input name="ListPaingType" type="radio" class="noborder" onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='none';ListPaing3.style.display=''" value="3" <%If ListPaingType=3 Then Response.Write "checked"%>>
		手动添加      </td>
    </tr>
	<tr id="ListPaing1" style="display:'<%If ListPaingType<>1 Then Response.Write "none"%>'">
      <td width="20%" class="b1_1"><strong><font color=blue>下页开始标记：</font></strong>
        <p>　</p><p>　</p>
      <strong><font color=blue>下页结束标记：</font></strong>      </td>
      <td width="75%" class="b1_1">
		<textarea name="LPsString" cols="49" rows="7"><%=LPsString%></textarea><br>
		<textarea name="LPoString" cols="49" rows="7"><%=LPoString%></textarea>      </td>
    </tr>
	<tr id="ListPaing12" style="display:'<%If ListPaingType<>1 Then Response.Write "none"%>'">
      <td width="20%" class="b1_1"><strong><font color=blue>索引分页重定向：</font></strong>      </td>
      <td width="75%" class="b1_1">
		<input name="ListPaingStr1" type="text" size="58" maxlength="200" value="<%=ListPaingStr1%>">      </td>
    </tr>
    <tr id="ListPaing2" style="display:'<%If ListPaingType<>2 Then Response.Write "none"%>'"> 
      <td width="20%" class="b1_1"><strong><font color=blue>批量生成：</font></strong></td>
      <td width="75%" class="b1_1">
        原字符串：<br>
		<input name="ListPaingStr2" type="text" size="58" maxlength="200" value="<%=ListPaingStr2%>"><br>
                格式：http://www.laoy.net/list.asp?page={$ID}<br>
                <br>
	    生成范围：<br>
		<input name="ListPaingID1" type="text" size="8" maxlength="200" value="<%=ListPaingID1%>"><span lang="en-us"> To </span><input name="ListPaingID2" type="text" size="8" maxlength="200" value="<%=ListPaingID2%>"><br>
               格式：只能是数字，可升序或者降序。      </td>
    </tr>
    <tr id="ListPaing3" style="display:'<%If ListPaingType<>3 Then Response.Write "none"%>'"> 
      <td width="20%" class="b1_1"><strong><font color=blue>手动添加：</font></strong>      </td>
      <td width="75%" class="b1_1">
      <textarea name="ListPaingStr3" cols="49" rows="7"><%=ListPaingStr3%></textarea><br>
      格式：输入一个网址后按回车，再输入下一个。      </td>
    </tr>

    <tr> 
      <td colspan="2" align="center" class="b1_1">
        <input name="ItemID" type="hidden" id="ItemID" value="<%=ItemID%>">
        <input name="Action" type="hidden" id="Action" value="SaveEdit">
        <input  type="button" name="button1" value="上&nbsp;一&nbsp;步" onClick="window.location.href='Admin_ItemModify.asp?ItemID=<%=ItemID%>'">
      <input  type="submit" name="Submit" value="下&nbsp;一&nbsp;步"></td>
    </tr>
</form>
</table>
<!--#include file="../Admin_Copy.asp"-->        
</body>         
</html>
<%End If%>
<%End Sub%>
<%
Sub SaveEdit
   ItemName=Trim(Request.Form("ItemName"))
   WebName=Trim(Request.Form("WebName"))
   WebUrl=Trim(Request.Form("WebUrl"))
   ClassID=Trim(Request.Form("ClassID"))
   SpecialID=Trim(Request.Form("SpecialID"))
   LoginType=Trim(Request.Form("LoginType"))
   LoginUrl=Trim(Request.Form("LoginUrl"))
   LoginPostUrl=Trim(Request.Form("LoginPostUrl"))
   LoginUser=Trim(Request.Form("LoginUser"))
   LoginPass=Trim(Request.Form("LoginPass"))
   LoginFalse=Trim(Request.Form("LoginFalse"))
   ItemDemo=Request.Form("ItemDemo")
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
      If ClassID="" or ClassID=0 Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>未指定栏目</li>"
      Else
         ClassID=Clng(ClassID)
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
         ErrMsg=ErrMsg & "<br><li>请选择网站登录类型</li>"
      Else
         LoginType=Clng(LoginType)
         If LoginType=1 Then
            If LoginUrl="" Or LoginPostUrl="" Or LoginUser="" Or LoginPass="" Or LoginFalse="" Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>网站登录信息不完整</li>"
            End If
         End If
      End If
   If FoundErr<>True Then
      SqlItem="Select top 1 *  from Item Where ItemID=" & ItemID
      Set RsItem=server.CreateObject("adodb.recordset")
      RsItem.Open SqlItem,ConnItem,2,3
      RsItem("ItemName")=ItemName
      RsItem("WebName")=WebName
      RsItem("WebUrl")=WebUrl
      RsItem("ClassID")=ClassID
      RsItem("LoginType")=LoginType
      If Logintype=1 Then
         RsItem("LoginUrl")=LoginUrl
         RsItem("LoginPostUrl")=LoginPostUrl
         RsItem("LoginUser")=LoginUser
         RsItem("LoginPass")=LoginPass
         RsItem("LoginFalse")=LoginFalse
      End If
      RsItem("ItemDemo")=ItemDemo
      RsItem.UpDate
      RsItem.Close
      Set RsItem=Nothing 
   End If
End Sub
%>
