<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
response.charset = "GB2312"
session.codepage = 936
session.timeout = 1440
server.scripttimeout = 9999
%>
<!--#include file="Config.asp"-->
<%
Dim sj:sj=0
Dim gl:gl=0
Dim fso
Dim laoystime : laoystime = timer()
Dim strobjectfso : strobjectfso = "scripting.filesystemobject"
Dim strobjectads : strobjectads = "Ado"&"db.Str"&"eam"
Dim strobjectxmlhttp : strobjectxmlhttp = "Mi" & "cr" & "os" & "of" & "t.X" & "M" & "LH" & "T" & "T" & "P"
	
	Const IsSqlDataBase = 0		'1ΪMsSQL,0ΪAccess
	Const IsDeBug = 0			'��������ģʽ�����Ե�ʱ������1���������е�ʱ������Ϊ0,�����������Ϣ�����ڰ�ȫ��
		
	dim conn,connstr,db
	dim tbname,SqlNowString
	
	tbname="Yao" 				'���ݱ�ǰ׺
	
	If IsSqlDataBase = 1 Then
		SqlNowString = "GetDate()"
	Else
		SqlNowString = "Now()"
	End If

	db=SitePath&"data/"&DataName
	on error resume next
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
	conn.Open connstr
	If Err Then
	err.Clear
	Set Conn = Nothing
	Response.Write "<div style='margin:100px;font-size:14px;text-align:center'>���ݿ����ӳ�������Inc/Config.Asp������ݿ����Ƽ�·���Ƿ���ȷ��</div>"
	Response.End
	End If
	
	LaoYID=Request.Cookies("Yao")("ID")
	LaoYName=Request.Cookies("Yao")("UserName")
	LaoYPass=Request.Cookies("Yao")("UserPass")
	LaoYRndPassword=Request.Cookies("Yao")("RndPassword")
	If LaoYID<>"" and LaoYName<>"" and LaoYPass<>"" and LaoYRndPassword<>"" then
	set rs4 = server.CreateObject ("adodb.recordset")
	sql="select * from "&tbname&"_User where id="& LaoYRequest(LaoYID) &" and PassWord='"&CheckStr(LaoYPass)&"' and RndPassword='"&CheckStr(LaoYRndPassword)&"' and UserName='"&CheckStr(LaoYName)&"'"
	on error resume next
	rs4.open sql,conn,1,1
	mymoney=rs4("UserMoney")
	username=rs4("UserName")
	useryn=rs4("yn")
	userqqopenid=rs4("qqopenid")
	LaoYdengji=rs4("usergroupid")
	UserPass=rs4("PassWord")
	UserRndPassword=rs4("RndPassword")
	rs4.close
	set rs4=nothing
		If UserPass<>LaoYPass or username<>LaoYName or UserRndPassword<>LaoYRndPassword Then
			LaoYID=""
			LaoYName=""
			LaoYPass=""
			LaoYRndPassword=""
		Else
			IsUser=1
		End if
	Else
		LaoYID=0
		LaoYName=""
		LaoYPass=""
		LaoYRndPassword=""
	End if
%>
<!--#include file="../Inc.asp"-->
<!--#include file="Function.asp"-->