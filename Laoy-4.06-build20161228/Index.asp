<!--#include file="Inc/conn.asp"--><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=SiteTitle%>|<%=SiteTitle2%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="keywords" content="<%=Sitekeywords%>" />
<meta name="description" content="<%=Sitedescription%>" />
<link href="<%=SitePath%>images/css<%=Css%>.css" type=text/css rel=stylesheet><%If rss=1 then%>
<link rel="alternate" title="����<%=SiteTitle%>(RSS 2.0)" href="http://<%=SiteUrl%><%=SitePath%>Rss/Rss.xml" type="application/rss+xml" /><%end if%>
<script type="text/javascript" src="<%=SitePath%>js/main.asp"></script>
<script type="text/javascript" src="<%=SitePath%>js/marquee.js"></script>
<style type="text/css">
#NEWLAOYBox{position:relative;}
#NEWLAOYNumID{position:absolute; bottom:5px; right:5px;}
#NEWLAOYNumID li{float:left;width:20px;height:20px;line-height:20px;text-align:center;background:url(images/sh-btn.gif) left no-repeat;color:#ffffff;cursor:pointer;margin-left:4px;overflow:hidden;}
#NEWLAOYNumID li:hover,#NEWLAOYNumID li.active{color:#FF7B11;background-position:right;}
#NEWLAOYContentID li{position:relative;width:250px;}
#NEWLAOYContentID .mask{FILTER:alpha(opacity=40);opacity:0.4;width:100%;height:65px;background-color:#000000;position:absolute;bottom:0;left:0;display:block;}
#NEWLAOYContentID .comt{float:left;width:230px;height:55px;position:absolute;left:0;padding:5px 0 0 10px;bottom:5px;font-size:12px;color:#ffffff;text-align:left;}
</style>
</head>
<body>
<%call spiderbot()%><div class="mwall">
<%=Head%>
<%=Menu%><div class="mw">
<div class="dh">
		<%=search%>�����ڵ�λ�ã�<a href="<%=SitePath%>"><%=SiteTitle%></a> >> ��ҳ
</div>
	<div id="ileft1">
        <div style="overflow:hidden;width:250px;height:270px;position:relative;clear:both;">
	<div id="NEWLAOYBox">
		<ul id="NEWLAOYContentID"><%
If IsHomeimg=0 then IsHomeimg=5
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top "&IsHomeimg&" ID,Title,Images from "&tbname&"_Article where yn = 0 and IsFlash=1 and Images<>'' order by ID desc"
rs1.open sql1,conn,1,3
If Not rs1.Eof Then 
do while not (rs1.eof or err) 
%>
			<li><a href="<%=apath(rs1("ID"),0)%>"><img border="0" src="<%=rs1("Images")%>" width="250" height="270"/></a><div class="mask"></div><div class="comt"><%=LoseHtml(replace(rs1("Title"),"'",""))%></div></li><%
  rs1.movenext
  loop
  end if
  rs1.close
  set rs1=nothing
  %>
		</ul>
	</div>
	<ul id="NEWLAOYNumID"><%for i=1 to IsHomeimg%><li><%=i%></li><%next%></ul>
</div>
<script type="text/javascript">
new Marquee(
{
	MSClassID : "NEWLAOYBox",
	ContentID : "NEWLAOYContentID",
	TabID	  : "NEWLAOYNumID",
	Direction : 2,
	Step	  : 0.5,
	Width	  : 250,
	Height	  : 270,
	Timer	  : 20,
	DelayTime : 3000,
	WaitTime  : 3000,
	ScrollStep: 250,
	SwitchType: 2,
	AutoStart : 1
})
</script><div id="clear"></div>
                <div id="ilist250" style="margin-top:4px;">
				<h6>�����Ƽ�</h6>
				<ul>
					<%Call ShowArticle(0,7,5,"",100,"IsHot=1","ID Desc",0,1,1)%>
				</ul>
				</div>
	</div>
	<div id="icenter1">
    	<div class="notice">
        	<div style="padding-top:2px;"><b>վ�ڹ��棺</b></div>
            <div id="laonotic" style="margin-top:5px;line-height:20px;height:20px; overflow:hidden;"> 
                <ul>
            		<%Call ShowArticle(NoticID,NoticNum,5,"��",100,"no","ID desc",0,1,0)%>
                </ul>
                <script type="text/javascript">
				<!--
				new Marquee("laonotic",0,5,400,20,50,3000,3000)
				//-->
				</script>
            </div>
		</div>
    	<div class="topnews">
			<%
            set rs1=server.createobject("ADODB.Recordset")
            sql1="select Top 1 ID,Title,Images,ArtDescription,TitleFontColor from "&tbname&"_Article where yn = 0 and IsTop = 1 and IsHot = 1 order by DateAndTime desc,ID Desc"
            rs1.open sql1,conn,1,3
            If Not rs1.Eof Then 
            do while not (rs1.eof or err) 
            %><h4><a href="<%=apath(rs1("ID"),0)%>"><%If rs1("TitleFontColor")<>"" then Response.Write("<font style=""color:"&rs1("TitleFontColor")&""">") End if%><%=rs1("Title")%><%If rs1("TitleFontColor")<>"" then Response.Write("</font>") End if%></a></h4>
              <div class="topjx">����<%=left(LoseHtml(rs1("ArtDescription")),100)%>����[<a href="<%=apath(rs1("ID"),0)%>">��ϸ</a>]</div>
            <%
              rs1.movenext
              loop
              end if
              rs1.close
              set rs1=nothing
            %>
		</div>
		<div id="clear"></div>
        <%Echo ShowAD(16)%>
        <div id="toplist">
			<ul>
            <%Call ShowArticle(0,10,5,"",100,"classid<>"&Noticid&"","DateAndTime desc,ID desc",1,1,0)%>
            </ul>
        </div>
	</div>
	<div id="iright">
    	<%Echo ShowAD(6)%>
	    <div class="nTableft">
      	<div class="TabTitleleft">
      		<ul id="myTab1">
					<li class="active" onMouseOver="nTabs(this,0);">��������</li>
        			<li class="normal" onMouseOver="nTabs(this,1);">һ������</li>
      		</ul>
    	</div>
		<div id="myTab1_Content0" style="clear:both;">
			<ul id="Artlist10num">
            <%Call ShowArticle(0,10,0,"",100,"no","Rnd(ID-timer())",0,1,1)%>
            </ul>
   		</div>
  		<div id="myTab1_Content1" class="none" style="clear:both;">
			<ul id="Artlist10num">
            <%Call ShowArticle(0,10,0,"",100,"datediff('d',DateAndTime,Now()) <= 30","Hits desc,ID desc",0,1,1)%>
            </ul>
		</div>
  		</div>
	</div>
</div>
<div id="clear"></div>
<%Echo ShowAD(14)%>
<div id="clear"></div>
<div class="mw">
    <div style="float:left;width:730px;">
<%
Dim rs,rs1
NoI=0
set rs=conn.execute("select ID,ClassName,IndexNum,Indeximg from "&tbname&"_Class Where IsIndex = 1 And link=0 order by num asc")
do while not rs.eof
NoI=NoI+1
Response.Write("		<div id=""ilist""")
If isEven(NoI)=0 then Response.Write(" style=""margin-left:6px;""") End if
Response.Write(">") & VbCrLf
Response.Write("			<h6><span><a href="""&cpath(rs(0),0)&""">more...</a></span><a href="""&cpath(rs(0),0)&""">"&rs(1)&"</a></h6>") & VbCrLf
	If rs(3)=1 then
	Response.Write("			<ul id=""indeximg"">") & VbCrLf
	Call ShowImgArticle(rs(0),1,100,"no","DateAndTime desc,ID desc",90,90,1)
	Response.Write("            </ul>") & VbCrLf
	End if
Response.Write("			<ul>") & VbCrLf
Call ShowArticle(rs(0),rs(2),5,"��",150,"no","DateAndTime desc,ID desc",0,1,0)
Response.Write("			</ul>") & VbCrLf
Response.Write("		</div>") & VbCrLf
If isEven(NoI)=0 then Response.Write("<div id=""clear""></div>") End if
	If NoI=4 then
	Response.Write("		<div id=""clear""></div>") & VbCrLf
	Echo ShowAD(12)
	Response.Write("		<div id=""clear""></div>") & VbCrLf
	elseIf NOI=8 then
	Response.Write("		<div id=""clear""></div>") & VbCrLf
	Echo ShowAD(13)
	Response.Write("		<div id=""clear""></div>") & VbCrLf
	End if
rs.movenext
loop
rs.close:set rs=nothing
%>
    </div>
	<div id="iright">
    	<%If indexpg>0 then%>
		<div class="nTableft" style="margin-bottom:4px;"> 
            <div class="TabTitleleft">
                <ul id="myTab11">
                        <li class="active" onMouseOver="nTabs(this,0);">��������</li>
                        <li class="normal" onMouseOver="nTabs(this,1);">��������</li>
                </ul>
            </div>
            <div id="myTab11_Content0" style="clear:both;padding:2px 0 1px 0">
<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top "&indexpg&" * from "&tbname&"_Guestbook where yn = 1 order by ID desc"
rs1.open sql1,conn,1,3
If Not rs1.Eof Then 
do while not (rs1.eof or err) 
If laoyvip then
	If rs1("yaoid")=0 then
	dispbbsid=rs1("id")
	else
	dispbbsid=rs1("yaoid")
	End if
End If
%>		<div class="igslist">
			<%=rs1("UserName")%>&nbsp;<%=day(rs1("AddTime"))%>��˵&nbsp;<a href="<%=IIf(laoyvip,"bbs/dispbbs.asp?id="&dispbbsid,sitepath&"Guestbook.asp")%>" target="_blank"><%=rs1("Title")%></a>
			<li><%=left(LoseHtml(rs1("Content")),40)%>..</li>
		</div>
  <%
  rs1.movenext
  loop
  else
  Response.Write("		<li>��û������!</li>")
  end if
  rs1.close
  set rs1=nothing
  %>
            </div>
            <div id="myTab11_Content1" class="none" style="clear:both;padding:2px 0 1px 0">
<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top "&indexpg&" * from "&tbname&"_Pl where yn = 1 order by ID desc"
rs1.open sql1,conn,1,3
If Not rs1.Eof Then 
do while not (rs1.eof or err)

set rsClass=server.createobject("adodb.recordset")
sql = "select id,title from "&tbname&"_Article where ID="&rs1("ArticleID")&""
rsClass.open sql,conn,1,1  
pTitle=rsClass("Title")
pid=rsclass("ID")
rsClass.close
set rsClass=nothing
%>		<div class="igslist">
			<%=rs1("memAuthor")%>&nbsp;��&nbsp;<a href="<%=apath(pid,0)%>" title="<%=pTitle%>"><%=left(pTitle,7)%></a>&nbsp;������
			<li><%=left(LoseHtml(ReplaceshowFace(rs1("memContent"))),40)%>..</li>
		</div>
  <%
  rs1.movenext
  loop
  else
  Response.Write("		<li>��û������!</li>")
  end if
  rs1.close
  set rs1=nothing
  %>
            </div>
  		</div>
        <%End if%>
        <%Echo ShowAD(15)%>
  <%If useroff=1 and indexuser>0 Then%><!-- ѡ���ʼ -->
  <div class="nTableft" style="margin-bottom:4px;">
    <!-- ���⿪ʼ -->
    <div class="TabTitleleft">
      <ul id="myTab3">
					<li class="active" onMouseOver="nTabs(this,0);">���¼����Ա</li>
        			<li class="normal" onMouseOver="nTabs(this,1);">��Ա<%=moneyname%>����</li>
      </ul>
    </div>
    <!-- ���ݿ�ʼ -->
    <div class="TabContent3">
      <div id="myTab3_Content0">
        	<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top "&indexuser&" * from "&tbname&"_User Where yn=1 order by ID desc"
rs1.open sql1,conn,1,3
NoI=0
If Not rs1.Eof Then 
do while not (rs1.eof or err) 
NoI=NoI+1
%>	
				  <tr>
				    <td width="65%"><%=NoI%>.<a href="<%=SitePath%>User/ShowUser.asp?ID=<%=rs1("ID")%>" target="_blank"><%=rs1("UserName")%></a></td>
				    <td width="20%"><%If rs1("Sex")=1 then Response.Write("��") else Response.Write("<font color=red>Ů</font>") end if%></td>
				    <td style="text-align:center;"><%=rs1("UserMoney")%></td>
				  </tr>	
  <%
  rs1.movenext
  loop
  else
  Response.Write("		<li>��û��!</li>")
  end if
  rs1.close
  set rs1=nothing
  %>
				  <tr>
				    <td colspan="3" style="text-align:right;padding-right:10px;"><a href="<%=SitePath%>User/UserList.asp?t=0">����.....</a></td>
			      </tr>
  			</table>
	  </div>
      <div id="myTab3_Content1" class="none">
        	<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top "&indexuser&" * from "&tbname&"_User Where yn=1 order by UserMoney desc,id desc"
rs1.open sql1,conn,1,3
NoI=0
If Not rs1.Eof Then 
do while not (rs1.eof or err) 
NoI=NoI+1
%>	
				  <tr>
				    <td width="65%"><%=NoI%>.<a href="<%=SitePath%>User/ShowUser.asp?ID=<%=rs1("ID")%>" target="_blank"><%=rs1("UserName")%></a></td>
				    <td width="20%"><%If rs1("Sex")=1 then Response.Write("��") else Response.Write("<font color=red>Ů</font>") end if%></td>
				    <td style="text-align:center;"><%=rs1("UserMoney")%></td>
				  </tr>	
  <%
  rs1.movenext
  loop
  else
  Response.Write("		<li>��û��!</li>")
  end if
  rs1.close
  set rs1=nothing
  %>
				  <tr>
				    <td colspan="3" style="text-align:right;padding-right:10px;"><a href="<%=SitePath%>User/UserList.asp?t=1">����.....</a></td>
			      </tr>
  			</table>
	  </div>
    </div>
  </div>
  <!-- ѡ����� --><%End if:If indexnum>0 then%>
  	<div id="ilist212" style="margin-bottom:4px;">
    	<h6>��������</h6>
		<ul>
			<%Call ShowArticle(0,indexnum,0,"",100,"no","Hits desc,ID desc",0,1,1)%>
		</ul>
    </div><%End if:If IsVote<>0 then%>
    <div id="ilist212" class="indexvote" style="padding-bottom:4px">
    	<h6>���ŵ���</h6>
		<ul>
			<%Call ShowVote(IsVote)%>
		</ul>
    </div><%End if%>       
    </div>
</div>
<div id="clear"></div>
<div class="mw">
	<div class="link"><%Call Link(0,0,0,1)%>��<a href="Link.asp">����...</a>����<a href="javascript:void(0)" onClick="document.getElementById('light').style.display='block';document.getElementById('fade').style.display='block';document.getElementById('fade').style.width = document.body.scrollWidth + 'px';document.getElementById('fade').style.height = document.body.scrollHeight + 'px';document.getElementById('light').style.left = (parseInt(document.body.scrollWidth) - 400) / 2 + 'px';">��������</a>��</div>
</div>
<div class="mw">
	<div class="link"><%Call Link(0,0,1,0)%></div>
</div>
<div id="light" class="white_content"> 
	<h5><span><a href="javascript:void(1)" onClick="document.getElementById('light').style.display='none';document.getElementById('fade').style.display='none'"><img src="<%=SitePath%>images/close.gif" alt="�ر�"/></a></span>��������</h5>
<iframe id="link" src="<%=SitePath%>Reglink.asp" width="100%" height="170" SCROLLING="NO" FRAMEBORDER="0"></iframe>
</div> 
<div id="fade" class="black_overlay"></div> 
<%=Copy%>
</div>
</body>
</html>