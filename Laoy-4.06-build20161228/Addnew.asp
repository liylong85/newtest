<!--#include file="Inc/conn.asp"-->
<%
If IsUser<>1 and IsUserP=0 then
		Response.Write(escape("只有会员才能发表评论!如果你是会员请登录后发表！"))
		Response.End
End if
Dim Author,Content,sResult,ArticleID
If IsUser=1 then
Author = LaoYName
else
Author = iparray(GetIP)
End if
Content =dvHTMLEncode(unescape(Request.Form("Content")))
ArticleID =LaoYRequest(unescape(Request.Form("ArticleID")))

	If ChkSB(Content)=false Then
		Response.Write(escape("请不要发布违法及广告信息!"))
		Response.End
	End if
	If session("postpltime")<>"" then
		posttime8=DateDiff("s",session("postpltime"),Now())
  		if posttime8<yaopostgetime then
		posttime9=yaopostgetime-posttime8
		Response.Write(escape("请不要连续发表。"))
		Response.End
  		end if
	end if
	
sResult = Author + Content
Conn.Execute("INSERT Into "&tbname&"_Pl(memAuthor,memContent,PostTime,ArticleID,IP,yn)  Values ('"&Author&"','"&Content&"','"&now()&"','"&ArticleID&"','"&GetIP&"',"&pingoff&")")
Session("postpltime")=Now()
'If Err Then 
   'Response.Write(escape("出现错误!"))
'Else
	If pingoff=0 then
		Response.Write(escape("发布评论成功,但是需要管理员审核后才能显示！"))
	else
		Response.Write(escape("发布评论成功!"))
	End if
'End If
%>