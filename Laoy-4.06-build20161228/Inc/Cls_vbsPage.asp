<%
Class Cls_vbsPage
	Private oConn		'���Ӷ���
	Private iPagesize	'ÿҳ��¼��
	Private sPageName	'��ַ��ҳ��������
	Private sDbType
	Private iRecType	'��¼����(>0Ϊ����ȡֵ�ٸ�����߹̶�ֵ,0ִ��count���ô�cookies,-1ִ��count������cookies)
	Private sJsUrl		'Cls_jsPage.js��·��
	Private sField		'�ֶ���
	Private sTable		'����
	Private sCondition	'����,����Ҫwhere
	Private sOrderBy	'����,����Ҫorder by,��Ҫasc����desc
	Private sPkey		'����,��д
	Private iRecCount

	'================================================================
	' Class_Initialize ��ĳ�ʼ��
	'================================================================
	Private Sub Class_Initialize
		iPageSize=10
		sPageName="Page"
		sDbType=IsSqlDataBase
		iRecType=0
		sJsUrl=""
		sField=" * "
	End Sub

	'================================================================
	' Conn �õ����ݿ����Ӷ���
	'================================================================
	Public Property Set Conn(ByRef Value)
		Set oConn=Value
	End Property

	'================================================================
	' PageSize ����ÿһҳ��¼����,Ĭ��10��¼
	'================================================================
	Public Property Let PageSize(ByVal intPageSize)
		iPageSize=CheckNum(intPageSize,0,0,iPageSize,0) 
	End Property

	'================================================================
	' PageName ��ַ��ҳ��������
	'================================================================
	Public Property Let PageName(ByVal strPageName)
		sPageName=IIf(Len(strPageName)<1,sPageName,strPageName)
	End Property

	'================================================================
	' DbType �õ����ݿ�����
	'================================================================ 
	Public Property Let DbType(ByVal strDbType)
		sDbType=UCase(IIf(Len(strDbType)<1,sDbType,strDbType))
	End Property

	'================================================================
	' RecType ȡ��¼����(>0Ϊ��ֵ���߹̶�ֵ,0ִ��count���ô�cookies,-1ִ��count������cookies����������)
	'================================================================
	Public Property Let RecType(ByVal intRecType)
		iRecType=CheckNum(intRecType,0,0,iRecType,0) 
	End Property

	'================================================================
	' JsUrl ȡ��Cls_jsPage.js��·��
	'================================================================
	Public Property Let JsUrl(ByVal strJsUrl)
		sJsUrl=strJsUrl
	End Property

	'================================================================
	' Pkey ȡ������
	'================================================================
	Public Property Let Pkey(ByVal strPkey)
		sPkey=strPkey
	End Property

	'================================================================
	' Field ȡ���ֶ���
	'================================================================
	Public Property Let Field(ByVal strField)
		sField=IIf(Len(strField)<1,sField,strField)
	End Property

	'================================================================
	' Table ȡ�ñ���
	'================================================================
	Public Property Let Table(ByVal strTable)
		sTable=strTable
	End Property

	'================================================================
	' Condition ȡ������
	'================================================================
	Public Property Let Condition(ByVal strCondition)
		s=strCondition
		sCondition=IIf(Len(s)>2," WHERE "&s,"")
	End Property

	'================================================================
	' OrderBy ȡ������
	'================================================================
	Public Property Let OrderBy(ByVal strOrderBy)
		s=strOrderBy
		sOrderBy=IIf(Len(s)>4," ORDER BY "&s,"")
	End Property

	'================================================================
	' RecCount ������¼����
	'================================================================
	Public Property Get RecCount()
		If iRecType>0 Then
			i=iRecType
		Elseif iRecType=0 Then
			i=CheckNum(Request.Cookies("ShowoPage")(sPageName),1,0,0,0)
			s=Trim(Request.Cookies("ShowoPage")("sCond"))
			IF i=0 OR sCondition<>s Then
				i=oConn.Execute("SELECT COUNT("&sPkey&") FROM "&sTable&" "&sCondition,0,1)(0)
				Response.Cookies("ShowoPage")(sPageName)=i
				Response.Cookies("ShowoPage")("sCond")=sCondition
			End If
		Else
			i=oConn.Execute("SELECT COUNT("&sPkey&") FROM "&sTable&" "&sCondition,0,1)(0)
		End If
		iRecCount=i
		RecCount=i
	End Property

	'================================================================
	' ResultSet ���ط�ҳ��ļ�¼��
	'================================================================
	Public Property Get ResultSet()
		s=Null
		'��¼����
		i=iRecCount
		'��ǰҳ
		If i>0 Then
			iPageCount=Abs(Int(-Abs(i/iPageSize)))'ҳ��
			iPageCurr=CheckNum(Request.QueryString(sPageName),1,1,1,iPageCount)'��ǰҳ
			Select Case sDbType
				Case 1 'sqlserver2000���ݿ�洢���̰�
					Set Rs=server.CreateObject("Adodb.RecordSet")
					Set Cm=Server.CreateObject("Adodb.Command")
					Cm.CommandType=4
					Cm.ActiveConnection=oConn
					Cm.CommandText="sp_laoy_classpage"
					Cm.parameters(1)=i
					Cm.parameters(2)=iPageCurr
					Cm.parameters(3)=iPageSize
					Cm.parameters(4)=sPkey
					Cm.parameters(5)=sField
					Cm.parameters(6)=sTable
					Cm.parameters(7)=Replace(sCondition," WHERE ","")
					Cm.parameters(8)=Replace(sOrderBy," ORDER BY ","")
					Rs.CursorLocation=3
					Rs.LockType=1
					Rs.Open Cm
				Case Else '�����������ԭʼ�ķ�������(ACͬ��)
					Set Rs = Server.CreateObject ("Adodb.RecordSet")
					ResultSet_Sql="SELECT "&sField&" FROM "&sTable&" "&sCondition&" "&sOrderBy
					Rs.Open ResultSet_Sql,oConn,1,1,&H0001
					Rs.AbsolutePosition=(iPageCurr-1)*iPageSize+1
			End Select
			s=Rs.GetRows(iPageSize)
			Rs.close
			Set Rs=Nothing
		End If
		ResultSet=s
	End Property

	'================================================================
	' Class_Terminate ��ע��
	'================================================================
	Private Sub Class_Terminate()
		If IsObject(oConn) Then oConn.Close:Set oConn=Nothing
	End Sub

	'================================================================
	' ����:����ַ�,�Ƿ�����Сֵ,�Ƿ������ֵ,��Сֵ(Ĭ������),���ֵ
	'================================================================
	Private Function CheckNum(ByVal strStr,ByVal blnMin,ByVal blnMax,ByVal intMin,ByVal intMax)
		Dim i,s,iMi,iMa
		s=Left(Trim(""&strStr),32):iMi=intMin:iMa=intMax
		If IsNumeric(s) Then
			i=CDbl(s)
			i=IIf(blnMin=1 And i<iMi,iMi,i)
			i=IIf(blnMax=1 And i>iMa,iMa,i)
		Else
			i=iMi
		End If
		CheckNum=i
	End Function

	'================================================================
	' ����ҳ����
	'================================================================
	Public Sub ShowPage()%>
		<Script Language="JavaScript" type="text/JavaScript" src="<%=sitepath%>js/Cls_jsPage.js"></Script>
		<Script Language="JavaScript" type="text/JavaScript">
		var s= new Cls_jsPage(<%=iRecCount%>,<%=iPageSize%>,2,"s"); 
		s.setPageSE("<%=sPageName%>=","");
		s.setPageInput("<%=sPageName%>");
		s.setUrl("");
		s.setPageFrist("��ҳ","��ҳ");
		s.setPagePrev("��ҳ","��һҳ");
		s.setPageNext("��ҳ","��һҳ");
		s.setPageLast("βҳ","βҳ");
		s.setPageText("{$PageNum}","{$PageNum}");
		s.setPageTextF(" {$PageTextF} "," {$PageTextF} ");
		s.setPageSelect("{$PageNum}","{$PageNum}");
		s.setPageCss("","","");
		s.setHtml("{$PageFrist}{$PagePrev}{$PageText}{$PageNext}{$PageLast}<li><span>��<font color='#009900'><b>{$RecCount}</b></font>��¼  {$PageSize}��/ҳ</span></li>");
		s.Write();
		</Script>
	<%End Sub

End Class%>