<!--#include file="Inc/conn.asp"-->
<!--#include file="Inc/Function_Page.asp"-->
<%
Dim KeyWord,stype,ID
KeyWord=server.URLEncode(request("KeyWord"))
stype=LaoYRequest(request("type"))
ID=LaoYRequest(request("ID"))

if request.QueryString("action")="search" then
	'dim keyword,laoy8
	'keyword = Server.UrlEncode(KeyWord)
	laoy8 = request.Form("bbs")
	Select case laoy8
		case "1"
			response.Redirect("Search.asp?KeyWord="&keyword)
		case "2"
			response.Redirect("Guestbook.asp?KeyWord="&keyword)
		case "3"
			response.Redirect("http://www.baidu.com/s?wd="&keyword&"")
		case "4"
			response.Redirect("http://www.google.com.hk/search?hl=zh-CN&q="&keyword)
		case "5"
			response.Redirect("http://www.youdao.com/search?q="&keyword)
		case "6"
			response.Redirect("http://one.cn.yahoo.com/s?p="&keyword)
	case else
		response.Redirect("Search.asp?KeyWord="&keyword)
	end select
end if
KeyWord = CheckStr(URLDecode(KeyWord))
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="keywords" content="<%=Sitekeywords%>" />
<meta name="description" content="<%=Sitedescription%>" />
<link href="<%=SitePath%>images/css<%=Css%>.css" type=text/css rel=stylesheet>
<script type="text/javascript" src="<%=SitePath%>js/main.asp"></script>
<title>����<%=KeyWord%>-<%=SiteTitle%></title>
</head>
<body>
<div class="mwall">
<%=Head%>
<%=Menu%><div class="mw">
	<div class="dh">
		<%=search%>�����ڵ�λ�ã�<a href="<%=sitepath%>">��ҳ</a> >> ����"<%=KeyWord%>"
    </div>
	<div id="nw_left">
		<div id="web2l">
        	<h6>վ������</h6>
			<div id="content">
				<ul id="listul">
<%
Set mypage=new xdownpage
mypage.getconn=conn

If KeyWord<>"" then
mypage.getsql=server.createobject("adodb.recordset")
If stype=1 or stype="" then
	If IsSqlDataBase = 1 Then
	mypage.getsql="select top 100 ID,Title,DateAndTime,Hits,IsTop,Images,TitleFontColor,Artdescription from "&tbname&"_Article Where yn = 0 and (Title like '%"&keyword&"%' or Artdescription like '%"&keyword&"%') order by DateAndTime desc"
	else
	mypage.getsql="select top 100 ID,Title,DateAndTime,Hits,IsTop,Images,TitleFontColor,Artdescription from "&tbname&"_Article Where yn = 0 and (InStr(1,LCase(Title),LCase('"&keyword&"'),0)<>0 or InStr(1,LCase(Artdescription),LCase('"&keyword&"'),0)<>0) order by DateAndTime desc"
	End if
elseIf stype=2 then
	mypage.getsql="select top 100 ID,Title,DateAndTime,Hits,IsTop,Images,TitleFontColor,Artdescription from "&tbname&"_Article Where yn = 0 and UserID = "&KeyWord&" order by DateAndTime desc"
End if

mypage.pagesize=""&artlistnum&""
set rs=mypage.getrs()
for i=1 to mypage.pagesize
    if not rs.eof then 
%>
				<li style="border-bottom:1px dashed #ccc;"><%if rs("Images")<>"" and artlist=0 then%><div style="float:left;margin:5px 5px 0 5px;"><a href="<%=apath(rs("ID"),0)%>" target="_blank"><img src="<%=rs("Images")%>" style="width:90px;height:90px;" alt="<%=rs("title")%>"/></a></div><%end if%><%If rs("IsTop")=1 then Response.Write("<font color=red>[��]</font>") end if%><a href="<%=apath(rs("ID"),0)%>" target="_blank"><%if rs("TitleFontColor")<>"" then Response.Write("<font style=""color:"&rs("TitleFontColor")&""">"&replacecolor(rs("Title"))&"</font>") else Response.Write(""&replacecolor(rs("Title"))&"") end if%></a>
				<span style="color:#AAA;font-size:12px;"><%=FormatDate(rs("DateAndTime"),11,1)%> �����<%=rs("Hits")%> ����:<%=Mydb("Select Count([ID]) From ["&tbname&"_Pl] Where yn=1 And ArticleID="&rs("id")&"",1)(0)%></span></li>
				<%If artlist=0 then%>
                <div class="box"<%if rs("Images")<>"" then Response.Write(" style=""height:65px;""") end if%>><%=left(LoseHtml(rs("Artdescription")),90)%>...</div><%End if%>
<%
        rs.movenext
    else
         exit for
    end if
next
%>
        		</ul>
			</div>
<div id="clear"></div>
<div id="page">
	<ul>
    <%=mypage.showpage()%>
    </ul>
</div>
<%else%>
������������ؼ��ֺ�����</ul></div><%end if%>
			</div>
		</div>
	<div id="nw_right">
		<%Echo ShowAD(3)%>
		<div id="web2r">
			<h5>ͼƬ����</h5>
			<ul class="topimg">
                <%Call ShowImgArticle(0,4,10,"no","DateAndTime desc,ID desc",60,60,1)%>
            </ul>
  		</div>
		<div id="web2r">
			<h5>�Ƽ�����</h5>
			<ul id="list10">
            	<%Call ShowArticle(0,10,5,"��",100,"IsHot=1","ID desc",0,0,0)%>
            </ul>
  		</div>
</div>
</div>
<%=Copy%></div>
</body>
</html>