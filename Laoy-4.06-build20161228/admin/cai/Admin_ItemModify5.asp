<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="../../inc/AutoKey.asp"-->
<%
Dim ItemID,Action
Dim RsItem,SqlItem,SqlF,RsF,FoundErr,ErrMsg
Dim LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,LoginResult,LoginData
Dim ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr,ChannelDir
Dim TsString,ToString,CsString,CoString,DateType,DsString,DoString,UpDateTime,AuthorType,AsString,AoString,AuthorStr,CopyFromType,FsString,FoString,CopyFromStr,KeyType,KsString,KoString,KeyStr,NewsPaingType,NPsString,NPoString,NewsPaingStr,NewsPaingHtml
Dim NewsPaingNext,NewsPaingNextCode,ContentTemp
Dim UrlTest,ListUrl,ListCode
Dim NewsUrl,NewsCode,NewsArrayCode,NewsArray
Dim Title,Content,Author,CopyFrom,Key
Dim Arr_Filters,Filteri,FilterStr
Dim UpDateType

Dim UploadFiles,strInstallDir,strChannelDir
strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/"))
strChannelDir="Test"
FoundErr=False

ItemID=Trim(Request("ItemID"))
Action=Trim(Request("Action"))

If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，项目ID不能为空</li>"
Else
   ItemID=Clng(ItemID)
End If

If Action="SaveEdit" And FoundErr<>True Then
   Call SaveEdit()
End If

If FoundErr<>True Then
   Call GetTest()
End If
If FoundErr<>True Then
   Call Main()
Else
   Call WriteErrMsg(ErrMsg)
End If
'关闭数据库链接
Call CloseConn()
Call CloseConnItem()
%>
<%Sub Main()%>
<html>
<head>
<title>采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../images/Admin_css.css">
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="30" class="b1_1"><a href="Admin_ItemAddNew.asp">添加项目</a> >> <a href="Admin_ItemModify.asp?ItemID=<%=ItemID%>">基本设置</a> >> <a href="Admin_ItemModify2.asp?ItemID=<%=ItemID%>">列表设置</a> >> <a href="Admin_ItemModify3.asp?ItemID=<%=ItemID%>">链接设置</a> >> <a href="Admin_ItemModify4.asp?ItemID=<%=ItemID%>">正文设置</a> >> <font color=red>采样测试</font> >> 属性设置 >> 完成</td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="admintable" >
  <tr>
    <td height="22" colspan="2" class="admintitle">编辑项目--采样测试</td>
  </tr>
  <tr>
    <td colspan="2" align="center" class="b1_1"><span lang="zh-cn"><%=Title%></span>　作者：<%=Author%>&nbsp;&nbsp;来源：<%=CopyFrom%>&nbsp;&nbsp;更新时间：<%=UpDateTime%></td>
  </tr>
  <tr>
    <td colspan="2" class="b1_1"><span lang="zh-cn"><span lang="zh-cn"><%=Content%></span></span></td>
  </tr>
  <tr>
    <td colspan="2" class="b1_1"><b>关键字：<%=key%></b></td>
  </tr>
  <tr>
    <td colspan="2" class="b1_1"><form method="post" action="Admin_ItemAttribute.asp" name="form1">
      <input name="Action" type="hidden" id="Action" value="SaveEdit">
              <input name="ItemID" type="hidden" id="ItemID" value="<%=ItemID%>">
              <input name="Cancel" type="button" id="Cancel" value="上&nbsp;一&nbsp;步" onClick="window.location.href='Admin_ItemModify4.asp?ItemID=<%=ItemID%>'">
<input  type="submit" name="Submit" value="下&nbsp;一&nbsp;步">
    </form></td>
  </tr>
</table>
<!--#include file="../Admin_Copy.asp"-->      
</body>         
</html>
<%End Sub%>

<%
'==================================================
'过程名：SaveEdit
'作  用：保存设置
'参  数：无
'==================================================
Sub SaveEdit

TsString=Request.Form("TsString")
ToString=Request.Form("ToString")
CsString=Request.Form("CsString")
CoString=Request.Form("CoString")

DateType=Trim(Request.Form("DateType"))
DsString=Request.Form("DsString")
DoString=Request.Form("DoString")

AuthorType=Trim(Request.Form("AuthorType"))
AsString=Request.Form("AsString")
AoString=Request.Form("AoString")
AuthorStr=Request.Form("AuthorStr")

CopyFromType=Trim(Request.Form("CopyFromType"))
FsString=Request.Form("FsString")
FoString=Request.Form("FoString")
CopyFromStr=Request.Form("CopyFromStr")

KeyType=Trim(Request.Form("KeyType"))
KsString=Request.Form("KsString")
KoString=Request.Form("KoString")
KeyStr=Request.Form("KeyStr")


NewsPaingType=Trim(Request.Form("NewsPaingType"))
NpsString=Request.Form("NpsString")
NpoString=Request.Form("NpoString")
NewsPaingStr=Request.Form("NewsPaingStr")
NewsPaingHtml=Request.Form("NewsPaingHtml")
if instr(dostring,"&nbsp;") then
response.write "ok"
end if
UrlTest=Trim(Request.Form("UrlTest"))

If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，项目ID不能为空</li>"
Else
   ItemID=Clng(ItemID)
End If
If UrlTest="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，数据传递时发生错误</li>"
Else
      NewsUrl=UrlTest
End If
If TsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>标题开始标记不能为空</li>"
End If
If ToString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>标题结束标记不能为空</li>" 
End If
If CsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>正文开始标记不能为空</li>"
End If
If CoString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>正文结束标记不能为空</li>" 
End If

If DateType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置时间类型</li>" 
Else
   DateType=Clng(DateType)
   If DateType=0 Then
   ElseIf DateType=1 Then
      If DsString="" or DoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请将时间的开始/结束标记填写完整</li>" 
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If
End If

If AuthorType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置作者类型</li>" 
Else
   AuthorType=Clng(AuthorType)
   If AuthorType=0 Then
   ElseIf AuthorType=1 Then
      If AsString="" or AoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请将作者开始/结束标记填写完整！</li>" 
      End If
   ElseIf AuthorType=2 Then
      If AuthorStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请指定作者</li>" 
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If 
End If

If CopyFromType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置来源类型</li>" 
Else
   CopyFromType=Clng(CopyFromType)
   If CopyFromType=0 Then
   ElseIf CopyFromType=1 Then
      If FsString="" or FoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请将来源开始/结束标记填写完整！</li>" 
      End If
   ElseIf CopyFromType=2 Then
      If CopyFromStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请指定来源</li>" 
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If 
End If

If KeyType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置关键字类型</li>" 
Else
   KeyType=Clng(KeyType)
   If KeyType=0 Then
   ElseIf KeyType=3 Then
   ElseIf KeyType=1 Then
      If KsString="" or KoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>关键字开始/结束标记不能为空</li>" 
      End If
   ElseIf KeyType=2 Then
      If KeyStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请指定关键字</li>" 
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If
End If

If NewsPaingType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置文章分页类型</li>"
Else
   NewsPaingType=Clng(NewsPaingType)
   If NewsPaingType=0 Then
   ElseIf NewsPaingType=1 Then
      If NPsString="" or NPoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>分页开始/结束标记不能为空</li>" 
      End If
      If NewsPaingStr<>"" And Len(NewsPaingStr)<15 Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>文章分页绝对链接设置不正确(至少15个字符)</li>" 
      End If
   ElseIf NewsPaingType=2 Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>暂不支持手动设置分页类型</li>" 
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If
End If

If FoundErr<>True Then
   SqlItem="Select * from Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,2,3
   RsItem("TsString")=TsString
   RsItem("ToString")=ToString
   RsItem("CsString")=CsString
   RsItem("CoString")=CoString

   RsItem("DateType")=DateType
   If DateType=1 Then
      RsItem("DsString")=DsString
      RsItem("DoString")=DoString
   End If

   RsItem("AuthorType")=AuthorType
   If AuthorType=1 Then
      RsItem("AsString")=AsString
      RsItem("AoString")=AoString
   ElseIf AuthorType=2 Then
      RsItem("AuthorStr")=AuthorStr
   End If

   RsItem("CopyFromType")=CopyFromType
   If CopyFromType=1 Then
      RsItem("FsString")=FsString
      RsItem("FoString")=FoString
   ElseIf CopyFromType=2 Then
      RsItem("CopyFromStr")=CopyFromStr
   End If

   RsItem("KeyType")=KeyType
   If KeyType=1 Then
      RsItem("KsString")=KsString
      RsItem("KoString")=KoString
   ElseIf KeyType=2 Then
      RsItem("KeyStr")=KeyStr
   End If

   RsItem("NewsPaingType")=NewsPaingType
   If NewsPaingType=1 Then
      RsItem("NPsString")=NPsString
      RsItem("NPoString")=NPoString
      RsItem("NewsPaingStr")=NewsPaingStr
      RsItem("NewsPaingHtml")=NewsPaingHtml
   ElseIf NewsPaingType=2 Then      
   End If
   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing
End If
End Sub

'==================================================
'过程名：GetTest
'作  用：采集测试
'参  数：无
'==================================================
Sub GetTest()
   SqlItem="Select * from Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If  RsItem.Eof  And  RsItem.Bof  Then
         FoundErr=True
      ErrMsg=ErrMsg  &  "<br><li>参数错误，找不到该项目</li>"
   Else
      LoginType=RsItem("LoginType")
      LoginUrl=RsItem("LoginUrl")
      LoginPostUrl=RsItem("LoginPostUrl")
      LoginUser=RsItem("LoginUser")
      LoginPass=RsItem("LoginPass")
      LoginFalse=RsItem("LoginFalse")
      ChannelDir=RsItem("ChannelDir")
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
      
      HsString=RsItem("HsString")
      HoString=RsItem("HoString")
      HttpUrlType=RsItem("HttpUrlType")
      HttpUrlStr=RsItem("HttpUrlStr")
      
      TsString=RsItem("TsString")
      ToString=RsItem("ToString")
      CsString=RsItem("CsString")
      CoString=RsItem("CoString")
      
      DateType=RsItem("DateType")
      DsString=RsItem("DsString")
      DoString=RsItem("DoString")

      AuthorType=RsItem("AuthorType")
      AsString=RsItem("AsString")
      AoString=RsItem("AoString")
      AuthorStr=RsItem("AuthorStr")

      CopyFromType=RsItem("CopyFromType")
      FsString=RsItem("FsString")
      FoString=RsItem("FoString")
      CopyFromStr=RsItem("CopyFromStr")

      KeyType=RsItem("KeyType")
      KsString=RsItem("KsString")
      KoString=RsItem("KoString")
      KeyStr=RsItem("KeyStr")

      NewsPaingType=RsItem("NewsPaingType")
      NPsString=RsItem("NPsString")
      NPoString=RsItem("NPoString")
      NewsPaingStr=RsItem("NewsPaingStr")
      NewsPaingHtml=RsItem("NewsPaingHtml")
      
      UpDateType=RsItem("UpDateType")
   End  If   
   RsItem.Close
   Set RsItem=Nothing

   If LoginType=1 Then
      If LoginUrl="" or LoginPostUrl="" or LoginUser="" Or LoginPass="" Or LoginFalse="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>您要采集的网站需要登录！请将登录信息填写完整</li>"
      End If
   End If
   If LsString="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>列表开始标记不能为空！</li>"
   End If
   If LoString="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>列表结束标记不能为空！</li>"
   End If
   If ListPaingType=0 Or ListPaingType=1 Then
      If ListStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>列表索引页不能为空！</li>"
      End If
      If ListPaingType=1 Then    
         If LPsString="" Or LPoString="" Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>索引分页开始、结束标记不能为空！</li>"
         End If
      End If      
      If  ListPaingStr1<>""  And  Len(ListPaingStr1)<15  Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>索引分页重定向设置不正确！</li>"
            End  IF
   ElseIf ListPaingType=2 Then
      If ListPaingStr2="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>批量生成原字符串不能为空！</li>"
      End If
      If IsNumeric(ListPaingID1)=False or IsNumeric(ListPaingID2)=False Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>批量生成的范围只能是数字！</li>"
      Else
         ListPaingID1=Clng(ListPaingID1)
         ListPaingID2=Clng(ListPaingID2)
         If ListPaingID1=0 And ListPaingID2=0 Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>批量生成的范围不正确！</li>"
         End If
      End If 
   ElseIf ListPaingType=3 Then
      If ListPaingStr3="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>索引分页不能为空！</li>"
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择索引分页类型</li>"
   End If
     If  HsString=""  or  HoString=""  Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>链接开始/结束标记不能为空！</li>"
      End  If


   If FoundErr<>True And  Action<>"SaveEdit"  Then
      Select Case ListPaingType
      Case 0,1
            ListUrl=ListStr
      Case 2
         ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1))
      Case 3
         If Instr(ListPaingStr3,"|")> 0 Then
            ListUrl=Left(ListPaingStr3,Instr(ListPaingStr3,"|")-1)
         Else
            ListUrl=ListPaingStr3
         End If
      End Select
   End If
   
      If  FoundErr<>True  And  Action<>"SaveEdit"  And  LoginType=1  Then
      LoginData=UrlEncoding(LoginUser & "&" & LoginPass)
      LoginResult=PostHttpPage(LoginUrl,LoginPostUrl,LoginData)
      If Instr(LoginResult,LoginFalse)>0 Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>登录网站时发生错误，请确认登录信息的正确性！</li>"
      End If
      End  If
   
   If  FoundErr<>True  And  Action<>"SaveEdit"  Then
         ListCode=GetHttpPage(ListUrl,ChannelDir)
         If ListCode<>"$False$" Then
            ListCode=GetBody(ListCode,LsString,LoString,False,False)
            If ListCode<>"$False$" Then
               NewsArrayCode=GetArray(ListCode,HsString,HoString,False,False)
               If NewsArrayCode<>"$False$" Then
                  If Instr(NewsArrayCode,"$Array$")>0 Then
                     NewsArray=Split(NewsArrayCode,"$Array$")
                     If HttpUrlType=1 Then
                        NewsUrl=Trim(Replace(HttpUrlStr,"{$ID}",NewsArray(0)))
                     Else
                        NewsUrl=Trim(DefiniteUrl(NewsArray(0),ListUrl))
                     End If
                  Else
                     FoundErr=True
                     ErrMsg=ErrMsg & "<br><li>只发现一个有效链接？：" & NewsArrayCode & "</li>"
                 End If
              Else
                 FoundErr=True
                 ErrMsg=ErrMsg & "<br><li>在获取链接列表时出错。</li>"
              End If   
           Else
               FoundErr=True
              ErrMsg=ErrMsg & "<br><li>在截取列表时发生错误。</li>"
           End If
        Else
            FoundErr=True
           ErrMsg=ErrMsg & "<br><li>在获取:" & ListUrl & "网页源码时发生错误。</li>"
        End If
     End If

If FoundErr<>True Then
   NewsCode=GetHttpPage(NewsUrl,ChannelDir)
   If NewsCode<>"$False$" Then
      Title=GetBody(NewsCode,TsString,ToString,False,False)
      Content=GetBody(NewsCode,CsString,CoString,False,False)
      If Title="$False$" or  Content="$False$"  Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>在截取文章标题/正文的时候发生错误：" & NewsUrl & "</li>"
      Else
         Title=FpHtmlEnCode(Title)
         Title=dvhtmlencode(Title)

         '文章分页
         If NewsPaingType=1 Then
            NewsPaingNext=GetPaing(NewsCode,NPsString,NPoString,False,False)
            Do While NewsPaingNext<>"$False$"
               If NewsPaingStr="" or Isnull(NewsPaingStr)=True Then
                  NewsPaingNext=DefiniteUrl(NewsPaingNext,NewsUrl)
               Else
                  NewsPaingNext=Replace(NewsPaingStr,"{$ID}",NewsPaingNext)
               End If
               If NewsPaingNext="" or NewsPaingNext="$False$" Then Exit Do
               NewsPaingNextCode=GetHttpPage(NewsPaingNext,ChannelDir)                  
               ContentTemp=GetBody(NewsPaingNextCode,CsString,CoString,False,False)
               If ContentTemp="$False$" Then
                  Exit Do
               Else
                  Content=Content & NewsPaingHtml & ContentTemp
                  NewsPaingNext=GetPaing(NewsPaingNextCode,NPsString,NPoString,False,False)
               End If
            Loop
         End If

         If UpDateType=0 Then
            UpDateTime=Now()
         ElseIf UpDateType=1 Then
            If DateType=0 then
               UpDateTime=Now()
            Else
               UpDateTime=GetBody(NewsCode,DsString,DoString,False,False)
               UpDateTime=FpHtmlEncode(UpDateTime)
               If IsDate(UpDateTime)=True Then
                  UpDateTime=CDate(UpDateTime)
               Else
                  UpDateTime=Now()
               End If
            End If
         ElseIf UpDateType=2 Then  
         Else
            UpDateTime=Now()
         End If

         '作者
         If AuthorType=1 Then
            Author=GetBody(NewsCode,AsString,AoString,False,False)
         ElseIf AuthorType=2 Then
            Author=AuthorStr
         End If
         If Author="$False$" Or Trim(Author)="" Then
            Author="佚名"
         Else
            Author=FpHtmlEnCode(Author)
         End If

         '来源
         If CopyFromType=1 Then
            CopyFrom=GetBody(NewsCode,FsString,FoString,False,False)
         ElseIf CopyFromType=2 Then
            CopyFrom=CopyFromStr
         End If
         If CopyFrom="$False$" Or Trim(CopyFrom)="" Then
            CopyFrom="不详"
         Else
            CopyFrom=FpHtmlEnCode(CopyFrom)
         End If

         If KeyType=0 Then
            Key=Title
			Key=CreateKeyWord(Key,2)
         ElseIf KeyType=3 Then
		 	Key=cn_split(Title,4)
         ElseIf KeyType=1 Then
            Key=GetBody(NewsCode,KsString,KoString,False,False)
            Key=FpHtmlEnCode(Key)
         ElseIf KeyType=2 Then
            Key=FpHtmlEnCode(KeyStr)
         End If
         If Key="$False$" Or Trim(Key)="" Then
            Key=Title
         End If
     End  If    
   Else
     FoundErr=True
     ErrMsg=ErrMsg & "<br><li>在获取源码时发生错误："& NewsUrl &"</li>"
   End If 
End If

If FoundErr<>True Then
   Call GetFilters
   Call Filters
   Content=ReplaceSaveRemoteFile(Content,strInstallDir,strChannelDir,False,NewsUrl)
End If

End Sub


'==================================================
'过程名：GetFilters
'作  用：提取过滤信息
'参  数：无
'==================================================
Sub GetFilters()
   SqlF ="Select * from Filters Where Flag=True And (PublicTf=True Or ItemID=" & ItemID & ") order by FilterID ASC"
   Set RSF=connItem.Execute(SqlF)
   If RsF.Eof And RsF.Bof Then
      Arr_Filters=""
   Else
      Arr_Filters=RsF.GetRows()
   End If
   RsF.Close
   Set RsF=Nothing
End Sub


'==================================================
'过程名：Filters
'作  用：过滤
'==================================================
Sub Filters()
If IsArray(Arr_Filters)=False Then
   Exit Sub
End if

   For Filteri=0 to Ubound(Arr_Filters,2)
      FilterStr=""
      If Arr_Filters(1,Filteri)=ItemID Or Arr_Filters(10,Filteri)=True Then
         If Arr_Filters(3,Filteri)=1 Then'标题过滤
            If Arr_Filters(4,Filteri)=1 Then
               Title=Replace(Title,Arr_Filters(5,Filteri),Arr_Filters(8,Filteri))
            ElseIf Arr_Filters(4,Filteri)=2 Then
               FilterStr=GetBody(Title,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Do While FilterStr<>"$False$"
                  Title=Replace(Title,FilterStr,Arr_Filters(8,Filteri))
                  FilterStr=GetBody(Title,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Loop
            End If
         ElseIf Arr_Filters(3,Filteri)=2 Then'正文过滤
            If Arr_Filters(4,Filteri)=1 Then
               Content=Replace(Content,Arr_Filters(5,Filteri),Arr_Filters(8,Filteri))
            ElseIf Arr_Filters(4,Filteri)=2 Then
               FilterStr=GetBody(Content,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Do While FilterStr<>"$False$"
                  Content=Replace(Content,FilterStr,Arr_Filters(8,Filteri))
                  FilterStr=GetBody(Content,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Loop
            End If
         End If
      End If
   Next
End Sub
%>