<% 
'Զ��ͼƬ��������
Const sFileExt="jpg|jpeg|gif|bmp|png"

'/////////////////////////////////////////////////////
'�� �ã��滻�ַ����е�Զ���ļ�Ϊ�����ļ�������Զ���ļ�
'�� ����
'      sHTML         : Ҫ�滻���ַ���
'      sSavePath     : �����ļ���·��
'      sExt          : ִ���滻����չ��
Function ReplaceRemoteUrl(sHTML, sSaveFilePath, sFileExt)
     Dim s_Content
     s_Content = sHTML
     If IsObjInstalled(strobjectxmlhttp) = False then
         ReplaceRemoteUrl = s_Content
         Exit Function
     End If
     
     Dim re, RemoteFile, RemoteFileurl,SaveFileName,SaveFileType,arrSaveFileNameS,arrSaveFileName,sSaveFilePaths
     Set re = new RegExp
     re.IgnoreCase = True
     re.Global = True
     re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(" & sFileExt & ")))"
     Set RemoteFile = re.Execute(s_Content)
     For Each RemoteFileurl in RemoteFile
		 arrSaveFileName = Split(RemoteFileurl,".")
  		 SaveFileType=arrSaveFileName(UBound(arrSaveFileName))
		 RanNum=Int(900*Rnd)+100
         arrSaveFileName = Year(Now()) & Right("0" & Month(Now()),2)&  Right("0" & Day(Now()),2) & Right("0" & Hour(Now()),2) & Right("0" & Minute(Now()),2) & Right("0" & Second(Now()),2) &ranNum&"."&SaveFileType
  		CreateFolder sSaveFilePath&"index.html"
         SaveFileName = sSaveFilePath & arrSaveFileName
         Call SaveRemoteFile(SaveFileName, RemoteFileurl)

		If IsAspJpeg=1 then
		Dim RV_img 
		RV_img=SaveFileName
		Call laoy_draw(RV_img)
		end if
		 
     s_Content = Replace(s_Content,RemoteFileurl,SaveFileName)
     Next
     ReplaceRemoteUrl = s_Content
End Function

'////////////////////////////////////////
'�� �ã�����Զ�̵��ļ�������
'�� ����LocalFileName ------ �����ļ���
'        RemoteFileUrl ------ Զ���ļ�URL
'����ֵ��True ----�ɹ�
'  False ----ʧ��
Sub SaveRemoteFile(s_LocalFileName,s_RemoteFileUrl)
     Dim Ads, Retrieval, GetRemoteData
     On Error Resume Next
     Set Retrieval = Server.CreateObject(strobjectxmlhttp)
     With Retrieval
         .Open "Get", s_RemoteFileUrl, False, "", ""
         .Send
         GetRemoteData = .ResponseBody
     End With
     Set Retrieval = Nothing
     Set Ads = Server.CreateObject(strobjectads)
     With Ads
         .Type = 1
         .Open
         .Write GetRemoteData
         .SaveToFile Server.MapPath(s_LocalFileName), 2
         .Cancel()
         .Close()
     End With
     Set Ads=nothing
End Sub

'////////////////////////////////////////
'�� �ã��������Ƿ��Ѿ���װ
'�� ����strClassString ----�����
'����ֵ��True ----�Ѿ���װ
'      False ----û�а�װ
Function IsObjInstalled(s_ClassString)
     On Error Resume Next
     IsObjInstalled = False
     Err = 0
     Dim xTestObj
     Set xTestObj = Server.CreateObject(s_ClassString)
     If 0 = Err Then IsObjInstalled = True
     Set xTestObj = Nothing
     Err = 0
End Function
%>