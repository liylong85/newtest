<%
'==================================================
'��������Admin_ShowChannel_Name
'��  �ã���ʾ��Ŀ����
'��  ����ClassID ------��ĿID
'==================================================
Sub Admin_ShowChannel_Name(ClassID)   
   Dim Sqlc,Rsc,TempStr
   ClassID=Clng(ClassID)
   Sqlc ="select top 1 ClassName from "&tbname&"_Class Where link = 0 and ID=" & ClassID   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.open Sqlc,Conn,1,1   
   If Rsc.Eof and Rsc.Bof then   
      TempStr="��ָ����Ŀ"   
   Else   
      TempStr=Rsc("ClassName")
   End if   
   Rsc.Close   
   Set Rsc=Nothing
   response.write TempStr   
End Sub  

'==================================================
'��������Admin_ShowChannel_Option
'��  �ã���ʾ��Ŀѡ��
'��  ����ClassID ------��ĿID
'==================================================
Sub Admin_ShowChannel_Option(ClassID)   
   Dim Sqlp,Rsp,TempStr,sqlpp,rspp
   Sqlp ="select * from "&tbname&"_Class Where link = 0 and TopID = 0 order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">��ѡ�����</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">�������ӷ���</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """" & Rsp("ID") & """" & "")
		 		If ClassID=Rsp("ID") Then
            	Response.Write(" selected")
         		End If
         Response.Write(">|-" & Rsp("ClassName") & "")
		 Response.Write("</option>" )  & VbCrLf
		    Sqlpp ="select * from "&tbname&"_Class Where TopID="&Rsp("ID")&" order by num"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof 
				Response.Write("<option value=" & """" & Rspp("ID") & """" & "")
				If ClassID=Rspp("ID") Then
            	Response.Write(" selected")
         		End If
         		Response.Write(">��|-" & Rspp("ClassName") & "")
				Response.Write("</option>" )  & VbCrLf
			Rspp.Movenext   
      		Loop
      Rsp.Movenext   
      Loop   
   End if  
End sub 


'==================================================
'��������Admin_ShowItem_Name
'��  �ã���ʾ��Ŀ����
'��  ����ItemID ------��ĿID
'==================================================
Sub Admin_ShowItem_Name(ItemID)   
   Dim Sqlc,Rsc,TempStr
   ItemID=Clng(ItemID)
   Sqlc ="select top 1 ItemName from Item Where ItemID=" & ItemID   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.open Sqlc,ConnItem,1,1   
   If Rsc.Eof and Rsc.Bof then   
      TempStr="��ָ����Ŀ"   
   Else   
      TempStr=Rsc("ItemName")
   End if   
   Rsc.Close   
   Set Rsc=Nothing
   Response.Write TempStr   
End Sub  


'==================================================
'��������Admin_ShowItem_Option
'��  �ã���ʾ��Ŀѡ��
'��  ����ItemID ------��ĿID
'==================================================
Sub Admin_ShowItem_Option(ItemID)   
   Dim SqlI,RsI,TempStr
   ItemID=Clng(ItemID)
   SqlI ="select ItemID,ItemName from Item order by ItemID desc"   
   Set RsI=server.CreateObject("adodb.recordset")   
   RsI.Open SqlI,ConnItem,1,1
   TempStr="<select Name=""ItemID"" ID=""ItemID"">"   
   If RsI.Eof and RsI.Bof Then
      TempStr=TempStr & "<option value=""0"">��������Ŀ</option>"   
   Else   
      TempStr=TempStr & "<option value=""0"">��ѡ����Ŀ</option>"
      Do while not RsI.Eof   
         TempStr=TempStr & "<option value=" & """" & RsI("ItemID") & """" & "" 
         If ItemID=RsI("ItemID") Then
            TempStr=TempStr & " Selected"
         End If
         TempStr=TempStr & ">" & RsI("ItemName")
         TempStr=TempStr & "</option>"  
      RsI.Movenext   
      Loop   
   End if
   RsI.Close   
   Set RsI=Nothing   
   TempStr=TempStr & "</select>"
   Response.Write TempStr   
End sub   

'==================================================
'��������GetHttpPage
'��  �ã���ȡ��ҳԴ��
'��  ����HttpUrl ------��ҳ��ַ
'==================================================
Function GetHttpPage(HttpUrl,laoy8)
   If IsNull(HttpUrl)=True Or Len(HttpUrl)<18 Or HttpUrl="$False$" Then
      GetHttpPage="$False$"
      Exit Function
   End If
   Dim Http
   on error resume next
   Set Http=server.createobject("Msxm"&"l2.X"&"MLHTTP")
   Http.open "GET",HttpUrl,False
   Http.Send()
   If Http.Readystate<>4 then
      Set Http=Nothing 
      GetHttpPage="$False$"
      Exit function
   End if
   GetHTTPPage=bytesToBSTR(Http.responseBody,laoy8)
   Set Http=Nothing
   If Err.number<>0 then
      Err.Clear
   End If
End Function

'==================================================
'��������BytesToBstr
'��  �ã�����ȡ��Դ��ת��Ϊ����
'��  ����Body ------Ҫת���ı���
'��  ����Cset ------Ҫת��������
'==================================================
Function BytesToBstr(Body,Cset)
   Dim Objstream
   Set Objstream = Server.CreateObject("adodb.stream")
   on error resume next
   objstream.Type = 1
   objstream.Mode =3
   objstream.Open
   objstream.Write body
   objstream.Position = 0
   objstream.Type = 2
   objstream.Charset = Cset
   BytesToBstr = objstream.ReadText 
   objstream.Close
   set objstream = nothing
End Function

'==================================================
'��������PostHttpPage
'��  �ã���¼
'==================================================
Function PostHttpPage(RefererUrl,PostUrl,PostData) 
    Dim xmlHttp 
    Dim RetStr      
    Set xmlHttp = CreateObject("Msxm"&"l2.X"&"MLHTTP")  
    xmlHttp.Open "POST", PostUrl, False
    XmlHTTP.setRequestHeader "Content-Length",Len(PostData) 
    xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlHttp.setRequestHeader "Referer", RefererUrl
    xmlHttp.Send PostData 
    If Err.Number <> 0 Then 
        Set xmlHttp=Nothing
        PostHttpPage = "$False$"
        Exit Function
    End If
    PostHttpPage=bytesToBSTR(xmlHttp.responseBody,"GB2312")
    Set xmlHttp = nothing
End Function 

'==================================================
'��������UrlEncoding
'��  �ã�ת������
'==================================================
Function UrlEncoding(DataStr)
    Dim StrReturn,Si,ThisChr,InnerCode,Hight8,Low8
    StrReturn = ""
    For Si = 1 To Len(DataStr)
        ThisChr = Mid(DataStr,Si,1)
        If Abs(Asc(ThisChr)) < &HFF Then
            StrReturn = StrReturn & ThisChr
        Else
            InnerCode = Asc(ThisChr)
            If InnerCode < 0 Then
               InnerCode = InnerCode + &H10000
            End If
            Hight8 = (InnerCode  And &HFF00)\ &HFF
            Low8 = InnerCode And &HFF
            StrReturn = StrReturn & "%" & Hex(Hight8) &  "%" & Hex(Low8)
        End If
    Next
    UrlEncoding = StrReturn
End Function

'==================================================
'��������GetBody
'��  �ã���ȡ�ַ���
'��  ����ConStr ------��Ҫ��ȡ���ַ���
'��  ����StartStr ------��ʼ�ַ���
'��  ����OverStr ------�����ַ���
'��  ����IncluL ------�Ƿ����StartStr
'��  ����IncluR ------�Ƿ����OverStr
'==================================================
Function GetBody(ConStr,StartStr,OverStr,IncluL,IncluR)
   If ConStr="$False$" or ConStr="" or IsNull(ConStr)=True Or StartStr="" or IsNull(StartStr)=True Or OverStr="" or IsNull(OverStr)=True Then
      GetBody="$False$"
      Exit Function
   End If
   Dim ConStrTemp
   Dim Start,Over
   ConStrTemp=Lcase(ConStr)
   StartStr=Lcase(StartStr)
   OverStr=Lcase(OverStr)
   Start = InStrB(1, ConStrTemp, StartStr, vbBinaryCompare)
   If Start<=0 then
      GetBody="$False$"
      Exit Function
   Else
      If IncluL=False Then
         Start=Start+LenB(StartStr)
      End If
   End If
   Over=InStrB(Start,ConStrTemp,OverStr,vbBinaryCompare)
   If Over<=0 Or Over<=Start then
      GetBody="$False$"
      Exit Function
   Else
      If IncluR=True Then
         Over=Over+LenB(OverStr)
      End If
   End If
   GetBody=MidB(ConStr,Start,Over-Start)
End Function


'==================================================
'��������GetArray
'��  �ã���ȡ���ӵ�ַ����$Array$�ָ�
'��  ����ConStr ------��ȡ��ַ��ԭ�ַ�
'��  ����StartStr ------��ʼ�ַ���
'��  ����OverStr ------�����ַ���
'��  ����IncluL ------�Ƿ����StartStr
'��  ����IncluR ------�Ƿ����OverStr
'==================================================
Function GetArray(Byval ConStr,StartStr,OverStr,IncluL,IncluR)
   If ConStr="$False$" or ConStr="" Or IsNull(ConStr)=True or StartStr="" Or OverStr="" or  IsNull(StartStr)=True Or IsNull(OverStr)=True Then
      GetArray="$False$"
      Exit Function
   End If
   Dim TempStr,TempStr2,objRegExp,Matches,Match
   TempStr=""
   Set objRegExp = New Regexp 
   objRegExp.IgnoreCase = True 
   objRegExp.Global = True
   objRegExp.Pattern = "("&StartStr&").+?("&OverStr&")"
   Set Matches =objRegExp.Execute(ConStr) 
   For Each Match in Matches
      TempStr=TempStr & "$Array$" & Match.Value
   Next 
   Set Matches=nothing

   If TempStr="" Then
      GetArray="$False$"
      Exit Function
   End If
   TempStr=Right(TempStr,Len(TempStr)-7)
   If IncluL=False then
      objRegExp.Pattern =StartStr
      TempStr=objRegExp.Replace(TempStr,"")
   End if
   If IncluR=False then
      objRegExp.Pattern =OverStr
      TempStr=objRegExp.Replace(TempStr,"")
   End if
   Set objRegExp=nothing
   Set Matches=nothing
   
   TempStr=Replace(TempStr,"""","")
   TempStr=Replace(TempStr,"'","")
   TempStr=Replace(TempStr," ","")
   TempStr=Replace(TempStr,"(","")
   TempStr=Replace(TempStr,")","")

   If TempStr="" then
      GetArray="$False$"
   Else
      GetArray=TempStr
   End if
End Function


'==================================================
'��������DefiniteUrl
'��  �ã�����Ե�ַת��Ϊ���Ե�ַ
'��  ����PrimitiveUrl ------Ҫת������Ե�ַ
'��  ����ConsultUrl ------��ǰ��ҳ��ַ
'==================================================
Function DefiniteUrl(Byval PrimitiveUrl,Byval ConsultUrl)
   Dim ConTemp,PriTemp,Pi,Ci,PriArray,ConArray
   If PrimitiveUrl="" or ConsultUrl="" or PrimitiveUrl="$False$" or ConsultUrl="$False$" Then
      DefiniteUrl="$False$"
      Exit Function
   End If
   If Left(Lcase(ConsultUrl),7)<>"http://" Then
      ConsultUrl= "http://" & ConsultUrl
   End If
   ConsultUrl=Replace(ConsultUrl,"\","/")
   ConsultUrl=Replace(ConsultUrl,"://",":\\")
   PrimitiveUrl=Replace(PrimitiveUrl,"\","/")

   If Right(ConsultUrl,1)<>"/" Then
      If Instr(ConsultUrl,"/")>0 Then
         If Instr(Right(ConsultUrl,Len(ConsultUrl)-InstrRev(ConsultUrl,"/")),".")>0 then   
         Else
            ConsultUrl=ConsultUrl & "/"
         End If
      Else
         ConsultUrl=ConsultUrl & "/"
      End If
   End If
   ConArray=Split(ConsultUrl,"/")

   If Left(LCase(PrimitiveUrl),7) = "http://" then
      DefiniteUrl=Replace(PrimitiveUrl,"://",":\\")
   ElseIf Left(PrimitiveUrl,1) = "/" Then
      DefiniteUrl=ConArray(0) & PrimitiveUrl
   ElseIf Left(PrimitiveUrl,2)="./" Then
      PrimitiveUrl=Right(PrimitiveUrl,Len(PrimitiveUrl)-2)
      If Right(ConsultUrl,1)="/" Then   
         DefiniteUrl=ConsultUrl & PrimitiveUrl
      Else
         DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & PrimitiveUrl
      End If
   ElseIf Left(PrimitiveUrl,3)="../" then
      Do While Left(PrimitiveUrl,3)="../"
         PrimitiveUrl=Right(PrimitiveUrl,Len(PrimitiveUrl)-3)
         Pi=Pi+1
      Loop            
      For Ci=0 to (Ubound(ConArray)-1-Pi)
         If DefiniteUrl<>"" Then
            DefiniteUrl=DefiniteUrl & "/" & ConArray(Ci)
         Else
            DefiniteUrl=ConArray(Ci)
         End If
      Next
      DefiniteUrl=DefiniteUrl & "/" & PrimitiveUrl
   Else
      If Instr(PrimitiveUrl,"/")>0 Then
         PriArray=Split(PrimitiveUrl,"/")
         If Instr(PriArray(0),".")>0 Then
            If Right(PrimitiveUrl,1)="/" Then
               DefiniteUrl="http:\\" & PrimitiveUrl
            Else
               If Instr(PriArray(Ubound(PriArray)-1),".")>0 Then 
                  DefiniteUrl="http:\\" & PrimitiveUrl
               Else
                  DefiniteUrl="http:\\" & PrimitiveUrl & "/"
               End If
            End If      
         Else
            If Right(ConsultUrl,1)="/" Then   
               DefiniteUrl=ConsultUrl & PrimitiveUrl
            Else
               DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & PrimitiveUrl
            End If
         End If
      Else
         If Instr(PrimitiveUrl,".")>0 Then
            If Right(ConsultUrl,1)="/" Then
               If right(LCase(PrimitiveUrl),3)=".cn" or right(LCase(PrimitiveUrl),3)="com" or right(LCase(PrimitiveUrl),3)="net" or right(LCase(PrimitiveUrl),3)="org" Then
                  DefiniteUrl="http:\\" & PrimitiveUrl & "/"
               Else
                  DefiniteUrl=ConsultUrl & PrimitiveUrl
               End If
            Else
               If right(LCase(PrimitiveUrl),3)=".cn" or right(LCase(PrimitiveUrl),3)="com" or right(LCase(PrimitiveUrl),3)="net" or right(LCase(PrimitiveUrl),3)="org" Then
                  DefiniteUrl="http:\\" & PrimitiveUrl & "/"
               Else
                  DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl
               End If
            End If
         Else
            If Right(ConsultUrl,1)="/" Then
               DefiniteUrl=ConsultUrl & PrimitiveUrl & "/"
            Else
               DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl & "/"
            End If         
         End If
      End If
   End If
   If Left(DefiniteUrl,1)="/" then
     DefiniteUrl=Right(DefiniteUrl,Len(DefiniteUrl)-1)
   End if
   If DefiniteUrl<>"" Then
      DefiniteUrl=Replace(DefiniteUrl,"//","/")
      DefiniteUrl=Replace(DefiniteUrl,":\\","://")
   Else
      DefiniteUrl="$False$"
   End If
End Function

'==================================================
'��������ReplaceSaveRemoteFile
'��  �ã��滻������Զ��ͼƬ
'��  ����ConStr ------ Ҫ�滻���ַ���
'��  ����SaveTf ------ �Ƿ񱣴��ļ���False�����棬True����
'��  ��: TistUrl------ ��ǰ��ҳ��ַ
'==================================================
Function ReplaceSaveRemoteFile(ConStr,strInstallDir,strChannelDir,SaveTf,TistUrl)
   If ConStr="$False$" or ConStr="" or strInstallDir="" Then
      ReplaceSaveRemoteFile=ConStr
      Exit Function
   End If
   Dim TempStr,TempStr2,TempStr3,Re,Matches,Match,Tempi,TempArray,TempArray2

   Set Re = New Regexp 
   Re.IgnoreCase = True 
   Re.Global = True
   Re.Pattern ="<img.+?[^\>]>"
   Set Matches =Re.Execute(ConStr) 
   For Each Match in Matches
      If TempStr<>"" then 
         TempStr=TempStr & "$Array$" & Match.Value
      Else
         TempStr=Match.Value
      End if
   Next
   If TempStr<>"" Then
      TempArray=Split(TempStr,"$Array$")
      TempStr=""
      For Tempi=0 To Ubound(TempArray)
         Re.Pattern ="src\s*=\s*.+?\.(gif|jpg|bmp|jpeg|psd|png|svg|dxf|wmf|tiff)"
         Set Matches =Re.Execute(TempArray(Tempi)) 
         For Each Match in Matches
            If TempStr<>"" then 
               TempStr=TempStr & "$Array$" & Match.Value
            Else
               TempStr=Match.Value
            End if
         Next
      Next
   End if
   If TempStr<>"" Then
      Re.Pattern ="src\s*=\s*"
      TempStr=Re.Replace(TempStr,"")
   End If
   Set Matches=nothing
   Set Re=nothing
   If TempStr="" or IsNull(TempStr)=True Then
      ReplaceSaveRemoteFile=ConStr
      Exit function
   End if
   TempStr=Replace(TempStr,"""","")
   TempStr=Replace(TempStr,"'","")
   TempStr=Replace(TempStr," ","")

   Dim RemoteFileurl,SavePath,PathTemp,DtNow,strFileName,strFileType,ArrSaveFileName,RanNum,Arr_Path
   DtNow=Now()
   If SaveTf=True then
 '***********************************
      SavePath= strChannelDir & ""&SitePath&""&SiteUp&"/" & year(DtNow) & right("0" & month(DtNow),2) & "/"
	  response.write "·����" & savepath & "<br>"
      Arr_Path=Split(SavePath,"/")
      PathTemp=""
      For Tempi=0 To Ubound(Arr_Path)
         If Tempi=0 Then
            PathTemp=Arr_Path(0) & "/"
         ElseIf Tempi=Ubound(Arr_Path) Then
            Exit For
         Else
            PathTemp=PathTemp & Arr_Path(Tempi) & "/"
         End If
         If CheckDir(PathTemp)=False Then
            If MakeNewsDir(PathTemp)=False Then
               SaveTf=False
               Exit For
            End If
         End If
      Next
   End If

   'ȥ���ظ�ͼƬ��ʼ
   TempArray=Split(TempStr,"$Array$")
   TempStr=""
   For Tempi=0 To Ubound(TempArray)
      If Instr(Lcase(TempStr),Lcase(TempArray(Tempi)))<1 Then
         TempStr=TempStr & "$Array$" & TempArray(Tempi)
      End If
   Next
   TempStr=Right(TempStr,Len(TempStr)-7)
   TempArray=Split(TempStr,"$Array$")
   'ȥ���ظ�ͼƬ����

   'ת�����ͼƬ��ַ��ʼ
   TempStr=""
   For Tempi=0 To Ubound(TempArray)
      TempStr=TempStr & "$Array$" & DefiniteUrl(TempArray(Tempi),TistUrl)
   Next
   TempStr=Right(TempStr,Len(TempStr)-7)
   TempStr=Replace(TempStr,Chr(0),"")
   TempArray2=Split(TempStr,"$Array$")
   TempStr=""
   'ת�����ͼƬ��ַ����

   'ͼƬ�滻/����
   Set Re = New Regexp
   Re.IgnoreCase = True 
   Re.Global = True

   For Tempi=0 To Ubound(TempArray2)
      RemoteFileUrl=TempArray2(Tempi)
      If RemoteFileUrl<>"$False$" And SaveTf=True Then'����ͼƬ
         ArrSaveFileName = Split(RemoteFileurl,".")
	 strFileType=Lcase(ArrSaveFileName(Ubound(ArrSaveFileName)))'�ļ�����
         If strFileType="asp" or strFileType="asa" or strFileType="aspx" or strFileType="cer" or strFileType="cdx" or strFileType="exe" or strFileType="rar" or strFileType="zip" then
            UploadFiles=""
            ReplaceSaveRemoteFile=ConStr
            Exit Function
         End If

         Randomize
         RanNum=Int(900*Rnd)+100
	 strFileName = year(DtNow) & right("0" & month(DtNow),2) & right("0" & day(DtNow),2) & right("0" & hour(DtNow),2) & right("0" & minute(DtNow),2) & right("0" & second(DtNow),2) & ranNum & "." & strFileType
         Re.Pattern =TempArray(Tempi)
	 If SaveRemoteFile(SavePath & strFileName,RemoteFileUrl)=True Then
'********************************
            PathTemp=SavePath & strFileName
            ConStr=Re.Replace(ConStr,PathTemp)
            Re.Pattern=strInstallDir & strChannelDir & "/"
            UploadFiles=UploadFiles & "|" & Re.Replace(SavePath &strFileName,"")
			
			If IsAspJpeg=1 then
			Dim RV_img 
			RV_img=SavePath&strFileName
			Call laoy_draw(RV_img)
			end if
         Else
            PathTemp=RemoteFileUrl
            ConStr=Re.Replace(ConStr,PathTemp)
            'UploadFiles=UploadFiles & "|" & RemoteFileUrl
         End If
      ElseIf RemoteFileurl<>"$False$" and SaveTf=False Then'������ͼƬ
         Re.Pattern =TempArray(Tempi)
         ConStr=Re.Replace(ConStr,RemoteFileUrl)
         UploadFiles=UploadFiles & "|" & RemoteFileUrl
      End If
   Next   
   Set Re=nothing
   If UploadFiles<>"" Then
      UploadFiles=Right(UploadFiles,Len(UploadFiles)-1)
   End If
   ReplaceSaveRemoteFile=ConStr
End function

'==================================================
'��������ReplaceSwfFile
'��  �ã���������·��
'��  ����ConStr ------ Ҫ�滻���ַ���
'��  ��: TistUrl------ ��ǰ��ҳ��ַ
'==================================================
Function ReplaceSwfFile(ConStr,TistUrl)
   If ConStr="$False$" or ConStr="" or TistUrl="" or TistUrl="$False$" Then
      ReplaceSwfFile=ConStr
      Exit Function
   End If
   Dim TempStr,TempStr2,TempStr3,Re,Matches,Match,Tempi,TempArray,TempArray2

   Set Re = New Regexp 
   Re.IgnoreCase = True 
   Re.Global = True
   Re.Pattern ="<object.+?[^\>]>"
   Set Matches =Re.Execute(ConStr) 
   For Each Match in Matches
      If TempStr<>"" then 
         TempStr=TempStr & "$Array$" & Match.Value
      Else
         TempStr=Match.Value
      End if
   Next
   If TempStr<>"" Then
      TempArray=Split(TempStr,"$Array$")
      TempStr=""
      For Tempi=0 To Ubound(TempArray)
         Re.Pattern ="value\s*=\s*.+?\.swf"
         Set Matches =Re.Execute(TempArray(Tempi)) 
         For Each Match in Matches
            If TempStr<>"" then 
               TempStr=TempStr & "$Array$" & Match.Value
            Else
               TempStr=Match.Value
            End if
         Next
      Next
   End if
   If TempStr<>"" Then
      Re.Pattern ="value\s*=\s*"
      TempStr=Re.Replace(TempStr,"")
   End If
   If TempStr="" or IsNull(TempStr)=True Then
      ReplaceSwfFile=ConStr
      Exit function
   End if
   TempStr=Replace(TempStr,"""","")
   TempStr=Replace(TempStr,"'","")
   TempStr=Replace(TempStr," ","")

   Set Matches=nothing
   Set Re=nothing

   'ȥ���ظ��ļ���ʼ
   TempArray=Split(TempStr,"$Array$")
   TempStr=""
   For Tempi=0 To Ubound(TempArray)
      If Instr(Lcase(TempStr),Lcase(TempArray(Tempi)))<1 Then
         TempStr=TempStr & "$Array$" & TempArray(Tempi)
      End If
   Next
   TempStr=Right(TempStr,Len(TempStr)-7)
   TempArray=Split(TempStr,"$Array$")
   'ȥ���ظ��ļ�����

   'ת����Ե�ַ��ʼ
   TempStr=""
   For Tempi=0 To Ubound(TempArray)
      TempStr=TempStr & "$Array$" & DefiniteUrl(TempArray(Tempi),TistUrl)
   Next
   TempStr=Right(TempStr,Len(TempStr)-7)
   TempStr=Replace(TempStr,Chr(0),"")
   TempArray2=Split(TempStr,"$Array$")
   TempStr=""
   'ת����Ե�ַ����

   '�滻
   Set Re = New Regexp
   Re.IgnoreCase = True 
   Re.Global = True
   For Tempi=0 To Ubound(TempArray2)
      RemoteFileUrl=TempArray2(Tempi)
      Re.Pattern =TempArray(Tempi)
      ConStr=Re.Replace(ConStr,RemoteFileUrl)
   Next   
   Set Re=nothing
   ReplaceSwfFile=ConStr
End function

'==================================================
'��������SaveRemoteFile
'��  �ã�����Զ�̵��ļ�������
'��  ����LocalFileName ------ �����ļ���
'��  ����RemoteFileUrl ------ Զ���ļ�URL
'==================================================
Function SaveRemoteFile(LocalFileName,RemoteFileUrl)
    SaveRemoteFile=True
	dim Ads,Retrieval,GetRemoteData
	Set Retrieval = Server.CreateObject("Microsof"&"t.X"&"MLHTTP")
	on error resume next
	With Retrieval
		.Open "Get", RemoteFileUrl, False, "", ""
		.Send
        If .Readystate<>4 then
            SaveRemoteFile=False
            Exit Function
        End If
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing
	Set Ads = Server.CreateObject("Adodb.Stream")
	With Ads
		.Type = 1
		.Open
		.Write GetRemoteData
		.SaveToFile server.MapPath(LocalFileName),2
		.Cancel()
		.Close()
	End With
	Set Ads=nothing
end Function

'==================================================
'��������FpHtmlEnCode
'��  �ã��������
'��  ����fString ------�ַ���
'==================================================
Function FpHtmlEnCode(fString)
   If IsNull(fString)=False or fString<>"" or fString<>"$False$" Then
       fString=nohtml(fString)
       fString=FilterJS(fString)
       fString = Replace(fString,"&nbsp;"," ")
       fString = Replace(fString,"&quot;","")
       fString = Replace(fString,"&#39;","")
       fString = replace(fString, ">", "")
       fString = replace(fString, "<", "")
       fString = Replace(fString, CHR(9), " ")'&nbsp;
       fString = Replace(fString, CHR(10), "")
       fString = Replace(fString, CHR(13), "")
       fString = Replace(fString, CHR(34), "")
       fString = Replace(fString, CHR(32), " ")'space
       fString = Replace(fString, CHR(39), "")
       fString = Replace(fString, CHR(10) & CHR(10),"")
       fString = Replace(fString, CHR(10)&CHR(13), "")
       fString=Trim(fString)
       FpHtmlEnCode=fString
   Else
       FpHtmlEnCode="$False$"
   End If
End Function

'==================================================
'��������GetPaing
'��  �ã���ȡ��ҳ
'==================================================
Function GetPaing(Byval ConStr,StartStr,OverStr,IncluL,IncluR)
If ConStr="$False$" or ConStr="" Or StartStr="" Or OverStr="" or IsNull(ConStr)=True or IsNull(StartStr)=True Or IsNull(OverStr)=True Then
   GetPaing="$False$"
   Exit Function
End If

Dim Start,Over,ConTemp,TempStr
TempStr=LCase(ConStr)
StartStr=LCase(StartStr)
OverStr=LCase(OverStr)
Over=Instr(1,TempStr,OverStr)
If Over<=0 Then
   GetPaing="$False$"
   Exit Function
Else
   If IncluR=True Then
      Over=Over+Len(OverStr)
   End If
End If
TempStr=Mid(TempStr,1,Over)
Start=InstrRev(TempStr,StartStr)
If IncluL=False Then
   Start=Start+Len(StartStr)
End If

If Start<=0 Or Start>=Over Then
   GetPaing="$False$"
   Exit Function
End If
ConTemp=Mid(ConStr,Start,Over-Start)

ConTemp=Trim(ConTemp)
ConTemp=Replace(ConTemp," ","")
ConTemp=Replace(ConTemp,",","")
ConTemp=Replace(ConTemp,"'","")
ConTemp=Replace(ConTemp,"""","")
ConTemp=Replace(ConTemp,">","")
ConTemp=Replace(ConTemp,"<","")
ConTemp=Replace(ConTemp,"&nbsp;","")
GetPaing=ConTemp
End Function

'==================================================
'��������ScriptHtml
'��  �ã�����html���
'��  ����ConStr ------ Ҫ���˵��ַ���
'==================================================
Function ScriptHtml(Byval ConStr,TagName,FType)
    Dim Re
    Set Re=new RegExp
    Re.IgnoreCase =true
    Re.Global=True
    Select Case FType
    Case 1
       Re.Pattern="<" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
    Case 2
       Re.Pattern="<" & TagName & "([^>])*>.*?</" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
    Case 3
       Re.Pattern="<" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
       Re.Pattern="</" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
    End Select
    ScriptHtml=ConStr
    Set Re=Nothing
End Function

Function CheckDir(byval FolderPath)
	dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(Server.MapPath(folderpath)) then
	'����
		CheckDir = True
	Else
	'������
		CheckDir = False
	End if
	Set fso = nothing
End Function
Function MakeNewsDir(byval foldername)
	dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder(Server.MapPath(foldername))
        If fso.FolderExists(Server.MapPath(foldername)) Then
           MakeNewsDir = True
        Else
           MakeNewsDir = False
        End If
	Set fso = nothing
End Function

'**************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'       False ----û�а�װ
'**************************************************
Function IsObjInstalled(strClassString)
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

'**************************************************
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'**************************************************
sub WriteErrMsg(ErrMsg)
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='../images/Admin_css.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='admintable' align=center>" & vbcrlf
	strErr=strErr & "  <tr><td class='admintitle' style=""text-align:center;"">������Ϣ</td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td style='padding:10px;'><b>��������Ŀ���ԭ��</b>" & ErrMsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td><a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'**************************************************
'��������WriteSucced
'��  �ã���ʾ�ɹ���ʾ��Ϣ
'��  ������
'**************************************************
sub WriteSucced(ErrMsg)
	dim strErr
	strErr=strErr & "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='../images/Admin_css.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=3 cellspacing=2 border=0 class='admintable' align=center>" & vbcrlf
	strErr=strErr & "  <tr><td class='admintitle' style=""text-align:center;"">��ϲ��</td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td align='center' style=""padding:20px;"">" & ErrMsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td align='center'><a href='Admin_ItemStart.asp'>&lt;&lt; ������ҳ</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'**************************************************
'��������ShowPage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sFileName  ----���ӵ�ַ
'       TotalNumber ----������
'       MaxPerPage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
'����ֵ������һҳ ��һҳ������Ϣ��HTML����
'**************************************************
function ShowPage(sFileName,TotalNumber,MaxPerPage,ShowTotal,ShowAllPages,strUnit)
	dim TotalPage,strTemp,strUrl,i

	if TotalNumber=0 or MaxPerPage=0 or isNull(MaxPerPage) then
		ShowPage=""
		exit function
	end if
	if totalnumber mod maxperpage=0 then
    	TotalPage= totalnumber \ maxperpage
  	else
    	TotalPage= totalnumber \ maxperpage+1
  	end if
	if CurrentPage>TotalPage then CurrentPage=TotalPage
		
  	strTemp= "<table align='center'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    	strTemp=strTemp & "��ҳ ��һҳ&nbsp;"
  	else
    	strTemp=strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
    	strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>��һҳ</a>&nbsp;"
  	end if

  	if CurrentPage>=TotalPage then
    	strTemp=strTemp & "��һҳ βҳ"
  	else
    	strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>��һҳ</a>&nbsp;"
    	strTemp=strTemp & "<a href='" & strUrl & "page=" & TotalPage & "'>βҳ</a>"
  	end if
   	strTemp=strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & TotalPage & "</strong>ҳ "
        strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;&nbsp;ת����<input type='text' name='page' size='3' maxlength='5' value='" & CurrentPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & strUrl & "page=" & "'+this.value;""'>ҳ"
	end if
	strTemp=strTemp & "</td></tr></table>"
	ShowPage=strTemp
end function
Dim fengyename
fengyename="["&"y"&"a"&"o"&"_"&"p"&"a"&"g"&"e"&"]"
'**************************************************
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
'**************************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function

'**************************************************
'��������CreateKeyWord
'��  �ã��ɸ������ַ������ɹؼ���
'��  ����Constr---Ҫ���ɹؼ��ֵ�ԭ�ַ���
'����ֵ�����ɵĹؼ���
'**************************************************
Function CreateKeyWord(byval Constr,Num)
   If Constr="" or IsNull(Constr)=True or Constr="$False$" Then
      CreateKeyWord="$False$"
      Exit Function
   End If
   If Num="" or IsNumeric(Num)=False Then
      Num=2
   End If
   Constr=Replace(Constr,CHR(32),"")
   Constr=Replace(Constr,CHR(9),"")
   Constr=Replace(Constr,"&nbsp;","")
   Constr=Replace(Constr," ","")
   Constr=Replace(Constr,"(","")
   Constr=Replace(Constr,")","")
   Constr=Replace(Constr,"<","")
   Constr=Replace(Constr,">","")
   Constr=Replace(Constr,"""","")
   Constr=Replace(Constr,"?","")
   Constr=Replace(Constr,"*","")
   Constr=Replace(Constr,"|","")
   Constr=Replace(Constr,",","")
   Constr=Replace(Constr,"��","")
   Constr=Replace(Constr,".","")
   Constr=Replace(Constr,"/","")
   Constr=Replace(Constr,"\","")
   Constr=Replace(Constr,"-","")
   Constr=Replace(Constr,"@","")
   Constr=Replace(Constr,"#","")
   Constr=Replace(Constr,"$","")
   Constr=Replace(Constr,"%","")
   Constr=Replace(Constr,"&","")
   Constr=Replace(Constr,"+","")
   Constr=Replace(Constr,":","")
   Constr=Replace(Constr,"��","")   
   Constr=Replace(Constr,"��","")
   Constr=Replace(Constr,"��","")
   Constr=Replace(Constr,"��","")         
   Dim i,ConstrTemp
   For i=1 To Len(Constr)
      ConstrTemp=ConstrTemp & Mid(Constr,i,Num) & "|"
   Next
   If Len(ConstrTemp)<11 Then
      ConstrTemp=ConstrTemp
   Else
      ConstrTemp=Left(ConstrTemp,11)
   End If
   CreateKeyWord=ConstrTemp
End Function

Function CheckUrl(strUrl)
   Dim Re
   Set Re=new RegExp
   Re.IgnoreCase =true
   Re.Global=True
   Re.Pattern="http://([\w-]+\.)+[\w-]+(/[\w-./?%&=]*)?"
   If Re.test(strUrl)=True Then
      CheckUrl=strUrl
   Else
      CheckUrl="$False$"
   End If
   Set Rs=Nothing
End Function
%>