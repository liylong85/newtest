<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
Dim Rs,Sql,SqlItem,RsItem,Action,FoundErr,ErrMsg
Dim HistrolyID,ItemID,ClassID,SpecialID,ArticleID,Title,CollecDate,NewsUrl,Result
Dim  Arr_Histroly,Arr_ArticleID,i_Arr,Del,Flag
Dim MaxPerPage,CurrentPage,AllPage,HistrolyNum,i_His
MaxPerPage=20
FoundErr=False
Del=Trim(Request("Del"))
Action=Trim(Request("Action"))
If Del="Del" Then
   Call DelHistroly()
End If
If FoundErr<>True Then
   Call Main()
else
   Call WriteErrMsg(ErrMsg)
End If
'关闭数据库链接
Call CloseConn()
Call CloseConnItem()
%>

<%Sub Main%>
<html>
<head>
<title>文章采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../images/Admin_css.css">
<style type="text/css">
.ButtonList {
	BORDER-RIGHT: #000000 2px solid; BORDER-TOP: #ffffff 2px solid; BORDER-LEFT: #ffffff 2px solid; CURSOR: default; BORDER-BOTTOM: #999999 2px solid; BACKGROUND-COLOR: #e6e6e6
}
</style>
<SCRIPT language=javascript>
function unselectall(thisform)
{
    if(thisform.chkAll.checked)
	{
		thisform.chkAll.checked = thisform.chkAll.checked&0;
    } 	
}

function CheckAll(thisform)
{
	for (var i=0;i<thisform.elements.length;i++)
    {
	var e = thisform.elements[i];
	if (e.Name != "chkAll"&&e.disabled!=true)
		e.checked = thisform.chkAll.checked;
    }
}
function submit(id){ 
document.getElementById(id).submit(); 
} 
</script>
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="30" class="b1_1"><a href="Admin_ItemHistroly.asp">管理首页</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="Admin_ItemHistroly.asp?Action=Succeed">成功记录</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="Admin_ItemHistroly.asp?Action=Failure">失败记录</a></td>
  </tr>
</table>
<%                             
Set RsItem=server.createobject("adodb.recordset")         
SqlItem="select * from Histroly"
If Action="Succeed"  Then
   SqlItem=SqlItem  &  " Where  Result=True"
   Flag="成 功 记 录"
ElseIf Action="Failure"  Then
   SqlItem=SqlItem  &  " Where  Result=False"
   Flag="失 败 记 录"
ElseIf Action="search"  Then
   SqlItem=SqlItem  &  " Where  ItemID = "&request("ItemID")&""
   Flag="搜索记录"
ElseIf  Action="LoseEf"  Then
   Flag="失 效 记 录"
   Set Rs=server.createobject("adodb.recordset")
   Sql="Select ID from "&tbname&"_Article"
   Rs.open  Sql,Conn,1,1
   If (Not Rs.Eof) And (Not Rs.Bof) Then
      Do While not rs.eof
         Arr_ArticleID=Arr_ArticleID & "," & CStr(rs("ID"))
      Rs.MoveNext
      Loop
   End  If
   Rs.Close
   Set Rs=Nothing
   If Arr_ArticleID<>"" Then
      Arr_ArticleID=0 & Arr_ArticleID
   Else
      Arr_ArticleID=0
   End If
   SqlItem=SqlItem  &  " Where ArticleID Not In("  &  Arr_ArticleID & ")"
Else
   Flag="所 有 记 录"
End  If
%>
<table width="100%" border="0" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable"> 
      <tr>         
      <td height="22" colspan="7" align="center" class=admintitle>历史记录管理</td>                 
     </tr>
   <form name="form1" method="POST" action="Admin_ItemHistroly.asp">         
     <tr style="padding: 0px 2px;">         
      <td width="5%" height="22" align="center" class=ButtonList>选择</td>                 
      <td width="25%" align="center" class=ButtonList>项目名称</td>         
      <td width="30%" align="center" class=ButtonList>文章标题</td>
      <td width="10%" height="22" align="center" class=ButtonList>栏目</td> 
      <td width="10%" align="center" class=ButtonList>来源</td>        
      <td width="10%" align="center" class=ButtonList>结果</td>
      <td width="10%" height="22" align="center" class=ButtonList>操作</td>         
     </tr>         
<%
If Request("page")<>"" then
    CurrentPage=Cint(Request("Page"))
Else
    CurrentPage=1
End if 
SqlItem=SqlItem  &  " order by HistrolyID DESC"
RsItem.open SqlItem,ConnItem,1,1
If (Not RsItem.Eof) and (Not RsItem.Bof) then
   RsItem.PageSize=MaxPerPage
   Allpage=RsItem.PageCount
   If Currentpage>Allpage Then Currentpage=1
   HistrolyNum=RsItem.RecordCount
   RsItem.MoveFirst
   RsItem.AbsolutePage=CurrentPage
   i_His=0
   Do While not RsItem.Eof
%>
    <tr onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;">          
      <td align="center" bgcolor="#f7f7f7"><input name="HistrolyID" type="checkbox" class="noborder" onClick="unselectall(this.form)" value="<%=RsItem("HistrolyID")%>"></td>               
      <td align="center" bgcolor="#f7f7f7">          
      <%Call Admin_ShowItem_Name(RsItem("ItemID"))%></td>         
      <td align="left" bgcolor="#f7f7f7"><%=RsItem("Title")%>      </td>  
      <td align="center" bgcolor="#f7f7f7"><%Call Admin_ShowChannel_Name(RsItem("ClassID"))%></td>
      <td align="center" bgcolor="#f7f7f7"><a href="<%=RsItem("NewsUrl")%>" target=_blank title=<%=RsItem("NewsUrl")%>>点击访问</a></td>     
      <td align="center" bgcolor="#f7f7f7">
      <%If RsItem("Result")=True Then
           Response.write "成功"
        ElseIf RsItem("Result")=False Then
           Response.Write "<font color=red>失败</font>"
        Else
           Response.Write "<font color=red>异常</font>"
        End If
      %></td>
      <td align="center" bgcolor="#f7f7f7">                           
      <a href="Admin_ItemHistroly.asp?Action=<%=Action%>&Del=Del&HistrolyID=<%=RsItem("HistrolyID")%>" onclick='return confirm("确定要删除此记录吗？");'>删除</a>      </td>         
    </tr>         
<%         
           i_His=i_His+1
           If i_His > MaxPerPage Then
              Exit Do
           End If
        RsItem.Movenext         
   Loop         
%>         
    <tr>          
      <td height="30" align="center" bgcolor="#f7f7f7">       
        <input name="Del" type="hidden" id="Del" value="Del">   
        <input name="Action" type="hidden" id="Action" value="<%=Action%>"> 
        <input name="chkAll" type="checkbox" class="noborder" id="chkAll" onclick=CheckAll(this.form) value="checkbox"></td>         
      <td height="30" colspan="7" bgcolor="#f7f7f7">全选
        <input type="submit" value="清除选中记录" name="DelFlag" onClick='return confirm("确定要清除所选记录吗？");'>
        <input type="submit" value="清除失败记录" name="DelFlag" onClick='return confirm("确定要清除所有失败记录吗？");'>
      <input type="submit" value="清空所有记录" name="DelFlag" onClick='return confirm("确定要清除所有记录吗？");'></td>
     </tr> </form> 
    <tr>          
      <td height="30" colspan=8 align="center" class="b1_1"><%
Response.Write ShowPage("Admin_ItemHistroly.asp?Action="& Action,HistrolyNum,MaxPerPage,True,True," 个记录")
%></td>         
    </tr>       
<%Else%>
<tr>
        <td colspan='9' align="center" class="b1_1"><br>系统中暂无历史记录！</td>
</tr> 
<%End  If%>       
<%         
RsItem.Close         
Set RsItem=nothing           
%>        
</table>  
<!--#include file="../Admin_Copy.asp"-->           
</body>         
</html>
<%End Sub%>
<%Sub DelHistroly
Dim DelFlag
DelFlag=Trim(Request("DelFlag"))
HistrolyID=Trim(Request("HistrolyID"))
If HistrolyID<>"" Then
   HistrolyID=Replace(HistrolyID," ","")
End If
If DelFlag="清除选中记录" Then
   If HistrolyID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择要删除的记录</li>"
   Else
      HistrolyID=Replace(HistrolyID," ","")
      SqlItem="Delete From [Histroly] Where HistrolyID in(" & HistrolyID & ")"
   End If
ElseIf DelFlag="清除失败记录" Then
   SqlItem="Delete From [Histroly] Where Result=False"
ElseIf DelFlag="清除失效记录" Then
   Set Rs=server.createobject("adodb.recordset")
   Sql="Select ArticleID From LZ8_Article Where Deleted=False"
   Rs.open Sql,Conn,1,1
   If (Not Rs.Eof) And (Not Rs.Bof) Then
      Do While not rs.eof
         Arr_ArticleID=Arr_ArticleID & "," & CStr(rs("ArticleID"))
      Rs.MoveNext
      Loop
   End  If
   Rs.Close
   Set Rs=Nothing
   If Arr_ArticleID<>"" Then
      Arr_ArticleID=0 & Arr_ArticleID
      SqlItem="Delete From [Histroly] Where ArticleID Not In(" & Arr_ArticleID & ")"
   Else
      SqlItem="Delete From [Histroly]"
   End If
ElseIf DelFlag="清空所有记录" Then
   SqlItem="Delete From [Histroly]"
Else
   If HistrolyID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择要删除的记录</li>"
   Else
      HistrolyID=Replace(HistrolyID," ","")
      SqlItem="Delete From [Histroly] Where HistrolyID In(" & HistrolyID & ")"
   End If
End if

If FoundErr<>True Then
   ConnItem.Execute(SqlItem)
End If
End Sub
%>