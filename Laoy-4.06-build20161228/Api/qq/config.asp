<script language="jscript" runat="server">
function getjson(str){
        try{
           eval("var jsonStr = (" + str + ")");
        }catch(ex){
           var jsonStr = null;
        }
        return jsonStr;
}
</script>
<%
'==================================
'=�� �� �ƣ�QqConnet
'=��    �ܣ�QQ��¼ For ASP
'=��    �ߣ��IFireFox�I
'=Q      Q: 63572063
'=��    �ڣ�2012-01-02
'==================================
'ת��ʱ�뱣���������ݣ���
Class QqConnet
    Private QQ_OAUTH_CONSUMER_KEY
    Private QQ_OAUTH_CONSUMER_SECRET
	Private QQ_CALLBACK_URL
	Private QQ_SCOPE
        
    Private Sub Class_Initialize      
        QQ_OAUTH_CONSUMER_KEY = "0"				'���APP ID
        QQ_OAUTH_CONSUMER_SECRET = "8fe40kd52a5e65d8209adf5c82e5467"		'���APP KEY
        QQ_CALLBACK_URL = "http://�������/api/qq/qqbind.asp"		'���ص�ַ������������
	QQ_SCOPE ="get_user_info,add_t,add_share,get_info,add_topic" 	'��Ȩ�� ���磺QQ_SCOPE=get_user_info,list_album,upload_pic,do_like,add_t 
                                                '������Ĭ������Խӿ�get_user_info������Ȩ��
                                                '���������Ȩ���������ֻ�����Ҫ�Ľӿ����ƣ���Ϊ��Ȩ��Խ�࣬�û�Խ���ܾܾ������κ���Ȩ��
    End Sub
    Property Get APP_ID()    
        APP_ID = QQ_OAUTH_CONSUMER_KEY    
    End Property

	'����Session("State")����.
	Public Function MakeRandNum()
		Randomize
		Dim width : width = 6 '���������,Ĭ��6λ
		width = 10 ^ (width - 1)
		MakeRandNum = Int((width*10 - width) * Rnd() + width)
	End Function
	
	Private Function CheckXml()
        Dim oxml,Getxmlhttp
        On Error Resume Next
        oxml=array("Microsoft.XMLHTTP","Msxml2.ServerXMLHTTP.6.0","Msxml2.ServerXMLHTTP.5.0","Msxml2.ServerXMLHTTP.4.0","Msxml2.ServerXMLHTTP.3.0","Msxml2.ServerXMLHTTP","Msxml2.XMLHTTP.6.0","Msxml2.XMLHTTP.5.0","Msxml2.XMLHTTP.4.0","Msxml2.XMLHTTP.3.0","Msxml2.XMLHTTP")
        For i=0 to ubound(oxml)
           Set Getxmlhttp = Server.CreateObject(oxml(i))
           If Err Then
              Err.Clear
              CheckXml = False
           Else
              CheckXml = oxml(i) :Exit Function
           End if
       Next
     End Function

	
	'Get��������url,��ȡ��������
	Private Function RequestUrl(url)
		Set XmlObj = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
		XmlObj.open "GET",url, false
		XmlObj.send
		If XmlObj.Readystate=4 Then
	       RequestUrl = XmlObj.responseText
	    Else
	       Response.Write("xmlhttp����ʱ��") 
		   Response.End()
	    End If
		Set XmlObj = nothing
	End Function
	
	'Post��������url,��ȡ��������
	Private Function RequestUrl_post(url,data)
		Set XmlObj = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
		XmlObj.open "POST", url, false
		XmlObj.setrequestheader "POST"," /t/add_t HTTP/1.1"
		XmlObj.setrequestheader "Host"," graph.qq.com "
		XmlObj.setrequestheader "content-length ",len(data)  
        XmlObj.setRequestHeader "Content-Type "," application/x-www-form-urlencoded "
		XmlObj.setrequestheader "Connection"," Keep-Alive"
        XmlObj.setrequestheader "Cache-Control"," no-cache"
        XmlObj.send(data)
		If XmlObj.Readystate=4 Then
	       RequestUrl_post = XmlObj.responseText
	    Else
	       Response.Write("xmlhttp����ʱ��") 
		   Response.End()
	    End If
		Set XmlObj = nothing
	End Function
	
	
	Private Function CheckData(data,str)
		If Instr(data,str)>0 Then
		   CheckData = True
		Else
		   CheckData = False
		End If
	End Function
	

	
	'���ɵ�¼��ַ
	Public Function GetAuthorization_Code()
		Dim url, params
		url = "https://graph.qq.com/oauth2.0/authorize"
		params = "client_id=" & QQ_OAUTH_CONSUMER_KEY
		params = params & "&redirect_uri=" & QQ_CALLBACK_URL
		params = params & "&response_type=code"
		params = params & "&scope="&QQ_SCOPE
		params = params & "&state="&Session("State")
		url = url & "?" & params
		GetAuthorization_Code = (url)
	End Function
	
	
	'��ȡ access_token
	Public Function GetAccess_Token()
		Dim url, params,Temp
		Url="https://graph.qq.com/oauth2.0/token"
	    params = "client_id=" & QQ_OAUTH_CONSUMER_KEY
		params = params & "&client_secret=" & QQ_OAUTH_CONSUMER_SECRET
		params = params & "&redirect_uri=" & QQ_CALLBACK_URL
		params = params & "&grant_type=authorization_code"
		params = params & "&code="&Session("Code")
		params = params & "&state="&Session("State")
		url = Url & "?" & params
		Temp=RequestUrl(url)
		If CheckData(Temp,"access_token=") = True Then
           GetAccess_Token=CutStr(Temp,"access_token=","&")
		Else
		   Response.Write("��ȡ access_token ʱ�������󣬴�����룺"&CutStr(Temp,"{""error"":",",")) 
		   Response.End()
		End If
	End Function
	
	'����Ƿ�Ϸ���¼��
	Public Function CheckLogin()
		Dim Code,mState
		Code=Trim(Request.QueryString("code"))
		If Code<>"" Then
			CheckLogin = True
			Session("Code")=Code
		Else
			CheckLogin = False
		End If
	End Function
	
	'��ȡopenid
	Public Function Getopenid()
		Dim url, params,Temp
		url = "https://graph.qq.com/oauth2.0/me"
		params = "access_token="&Session("Access_Token")
		url = Url & "?" & params
		Temp=RequestUrl(url)
		If Instr(Temp,"openid")>0 Then
		   set obj = getjson(CutStr(Temp,"(",")"))
		   if isobject(obj) Then
		       Getopenid=obj.openid
		   End If
		  set obj = Nothing
		Else
		   
		   set obj = getjson(CutStr(Temp,"(",")"))
		   if isobject(obj) Then
		       ret = obj.error
			   msg = obj.error_description
		   End If
		  set obj = Nothing
		    Response.Write("��ȡ openid ʱ�������󣬴�����룺"&ret&" , ����������"&msg) 
		   Response.End()
		End If
	End Function
		
	'��ȡ�û���Ϣ,�õ�һ��json��ʽ���ַ���
	Public Function GetUserInfo()
		Dim url, params, result
		url = "https://graph.qq.com/user/get_user_info"
		params = "oauth_consumer_key=" & QQ_OAUTH_CONSUMER_KEY
		params = params & "&access_token=" & Session("Access_Token")
		params = params & "&openid=" & Session("Openid")
		url = url & "?" & params
		Temp = RequestUrl(url)
		If CheckData(Temp,"nickname") = False Then
		    set obj = getjson(Temp)
		   if isobject(obj) Then
		       ret = obj.ret
			   msg = obj.msg
		   End If
		  set obj = Nothing
		   Response.Write("��ȡ�û���Ϣʱ�������󣬴�����룺"&ret&" , ����������"&msg) 
		   Response.End()
		End If
		GetUserInfo = Temp
	End Function
	
	'��ȡ��Ѷ΢����¼�û����û�����,�õ�һ��json��ʽ���ַ���
	Public Function Get_Info()
		Dim url, params, result
		url = "https://graph.qq.com/user/get_info"
		params = "oauth_consumer_key=" & QQ_OAUTH_CONSUMER_KEY
		params = params & "&access_token=" & Session("Access_Token")
		params = params & "&openid=" & Session("Openid")
		params = params & "&format=json"
		url = url & "?" & params
		Get_Info = RequestUrl(url)
	End Function

	
	'��ȡ�û�����,�Ա�,��json�ַ������ȡ����ַ�
	Public Function GetUserName(json)
	    Dim nickname,sex,obj
		set obj = getjson(json)
		   if isobject(obj) Then
		       nickname = obj.nickname
			   sex = obj.gender
		   End If
		  set obj = Nothing
	    GetUserName = Array(nickname,sex)
	End Function
	
	Public Function CutStr(data,s_str,e_str)
	    If Instr(data,s_str)>0 and Instr(data,e_str)>0 Then
		   CutStr = Split(data,s_str)(1)
		   CutStr = Split(CutStr,e_str)(0)
		Else
		   CutStr = ""
		End If
	End Function
	
End Class
%>