<!--#include file="Inc/conn.asp"-->
<%

dim classID,NclassID,NclassID1,showNclass,kind,dateNum,maxLen,listNum,bullet
dim hitColor,new_color,old_color

topType = CheckStr(Request("topType"))
If CheckStr(Request("ClassNo")) <> "" then
ClassNo = split(CheckStr(Request("ClassNo")),"|")
on error resume next
NClassID = LaoYRequest(ClassNo(0))
NClassID1 = LaoYRequest(ClassNo(1))
End if

num = LaoYRequest(request.querystring("num"))
maxlen = LaoYRequest(Request.querystring("maxlen"))
showdate = LaoYRequest(Request("showdate"))
showhits = LaoYRequest(Request("showhits"))
showClass = LaoYRequest(Request("showClass"))

bullet="·"			 '标题前的图片或符号
hitColor="#FF0000"   '点击数的颜色
new_color="#FF0000"  '新文章日期的颜色
old_color="#999999"  '旧文章日期的颜色


dim rs,sql,str,topic

set rs=server.createObject("Adodb.recordset")
sql = "Select top "& num &" ID,Title,TitleFontColor,Author,ClassID,DateAndTime,Hits,IsTop,IsHot from "&tbname&"_Article Where yn = 0"

	If NclassID<>"" and NclassID1="" then
		If Yao_MyID(NclassID)="0" then
			SQL=SQL&" and ClassID="&NclassID&""
		else
			MyID = Replace(""&Yao_MyID(NclassID)&"","|",",")
			SQL=SQL&" and ClassID in ("&MyID&")"
		End if
	elseif NclassID<>"" and NclassID1<>"" then
		MyID = Replace(""&CheckStr(Request("ClassNo"))&"","|",",")
		SQL=SQL&" and ClassID in ("&MyID&")"
	End if
	
select case topType
	case "new" sql=sql&" order by DateAndTime desc,ID desc"
	case "hot" sql=sql&" order by hits desc,ID desc"
	case "IsHot" sql=sql&" and IsHot = 1 order by ID desc"
end select

set rs = conn.execute(sql)
if rs.bof and rs.eof then 
str=str+"没有符合条件的文章"
else

rs.movefirst
do while not rs.eof
	topic=Left(LoseHtml(rs("title")),maxlen)
	topic=replace(server.HTMLencode(topic)," ","&nbsp;")
	topic=replace(topic,"'","&nbsp;")
	str=str+bullet
	if showClass = 1 then
		str=str+"[<a href='http://"&SiteUrl&cpath(rs("classid"),0)&"' target='_blank'>"&getclass(rs("ClassID"),"classname")&"</a>] "
	end if
	str=str+"<a href='http://"&SiteUrl&apath(rs("ID"),0)&"' target='_blank'  title='"&replace(replace(server.HTMLencode(rs("title"))," ","&nbsp;"),"'","&nbsp;")&"') >"+Topic+"</a>"
	if showdate <> 0 then
	str=str & FormatDate(rs("DateAndTime"),showdate,0)
	end if
	if showhits = 1 then
		str=str&"浏览：<font color="& hitcolor &">"& rs("hits") &"</font>"
	end if
	str=str & "<br />"
	rs.moveNext
loop
end if
rs.close : conn.close
set rs=nothing : set conn=nothing

response.write "document.write ("&Chr(34)&str&Chr(34)&");"
%>