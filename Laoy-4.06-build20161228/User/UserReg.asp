<!--#include file="../Inc/conn.asp"-->
<%
If LaoYID>0 then
Response.Redirect ""&SitePath&""
End if
If useroff=0 then Call Alert("��վĿǰ�Ѿ��رջ�Ա����","../")
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��Աע��-<%=SiteTitle%></title>
<meta name="keywords" content="<%=Sitekeywords%>" />
<meta name="description" content="<%=SiteTitle%>�û�ע��" />
<link href="<%=SitePath%>images/css<%=Css%>.css" type=text/css rel=stylesheet>
<script type="text/javascript" src="<%=SitePath%>js/main.asp"></script>
<script src="<%=SitePath%>js/Ajax.js" language="javascript"></script>
</head>
<body>
<div class="mwall">
<%=Head%>
<%=Menu%><div class="mw">
	<div class="dh">
		<%=search%>�����ڵ�λ�ã�<a href="<%=SitePath%>Index.asp">��ҳ</a> >> ��Աע��
    </div>
	<div id="nw_left">
		<div id="web2l">
        	<h6>��Աע��</h6>
			<div id="content">
			  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="30"><b>ע��</b>��<font color="#FF0000">*</font>Ϊ������</td>
                </tr>
              </table>
              <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-bottom:176px;">
                <form action="RegPost.asp" method="post" name="UserReg" id="UserReg" onSubmit="return user_from();" >
                  <tr>
                    <td width="30%" height="30">�û�����</td>
                    <td><input name="UserName"  class="borderall" value="" size="18" maxlength="15" style="width:140px;" onBlur="CheckName();" onChange="CheckName();" />
                        <font color="#FF0000">*</font></td>
                  </tr>
                  <tr>
                    <td height="30">&nbsp;</td>
                    <td><span id="enter_name">����������</span></td>
                  </tr>
                  <tr>
                    <td height="30">����(����6λ)��<br /><span id="enter_pwd"></span>                    </td>
                    <td><input name="UserPassword" type="password"  class="borderall" size="18" maxlength="12" style="width:140px;"/>
                        <font color="#FF0000">*</font></td>
                  </tr>
                  <tr>
                    <td height="30">ȷ������(����6λ)��<br /><span id="enter_repwd"></span></td>
                    <td><input name="PwdConfirm" type="password"  class="borderall" size="18" maxlength="12" style="width:140px;" />
                        <font color="#FF0000">*</font></td>
                  </tr>
                  <tr>
                    <td height="30">�Ա�</td>
                    <td><input name="UserSex" type="radio" value="1" checked="checked" />
                      ��
                      <input name="UserSex" type="radio" value="0" />
                      Ů</td>
                  </tr>
                  <tr>
                    <td height="30" class="td">�������ڣ�</td>
                    <td height="30" class="td"><select size="1" name="year" maxlength="4">
                        <option value="">��ѡ��</option>
                        <script language="JavaScript" type="text/javascript">writeOption(1940,<%=year(Now)-8%>);</script>
                      </select>
                      ��
                      <select size="1" name="month" maxlength="2">
                        <option value="">��ѡ��</option>
                        <script language="JavaScript" type="text/javascript">writeOption(1,12);</script>
                      </select>
                      ��
                      <select size="1" name="day" maxlength="2">
                        <option value="">��ѡ��</option>
                        <script language="JavaScript" type="text/javascript">writeOption(1,31);</script>
                      </select>
                      �ա�<font color="#FF0000">*</font></td>
                  </tr>
                  <tr>
                    <td height="30">Email��ַ��</td>
                    <td><input name="UserEmail"  class="borderall" size="25" maxlength="50"  onchange="UserEmail_enter();" onKeyUp="UserEmail_enter();" onBlur="UserEmail_enter();" />
                        <font color="#FF0000">*</font>��<span id="enter_mail"></span></td>
                  </tr>
                  <tr>
                    <td height="30">����(ʡ/��)��</td>
                    <td colspan="2"><select onChange="setcity();" name='province'>
                        <option value=''>��ѡ��ʡ��</option>
                        <option value="�㶫">�㶫</option>
                        <option value="����">����</option>
                        <option value="����">����</option>
                        <option value="����">����</option>
                        <option value="����">����</option>
                        <option value="����">����</option>
                        <option value="����">����</option>
                        <option value="����">����</option>
                        <option value="�ӱ�">�ӱ�</option>
                        <option value="������">������</option>
                        <option value="����">����</option>
                        <option value="���">���</option>
                        <option value="����">����</option>
                        <option value="����">����</option>
                        <option value="����">����</option>
                        <option value="����">����</option>
                        <option value="����">����</option>
                        <option value="����">����</option>
                        <option value="����">����</option>
                        <option value="���ɹ�">���ɹ�</option>
                        <option value="����">����</option>
                        <option value="�ຣ">�ຣ</option>
                        <option value="ɽ��">ɽ��</option>
                        <option value="�Ϻ�">�Ϻ�</option>
                        <option value="ɽ��">ɽ��</option>
                        <option value="����">����</option>
                        <option value="�Ĵ�">�Ĵ�</option>
                        <option value="����">����</option>
                        <option value="̨��">̨��</option>
                        <option value="���">���</option>
                        <option value="�½�">�½�</option>
                        <option value="����">����</option>
                        <option value="����">����</option>
                        <option value="�㽭">�㽭</option>
                        <option value="����">����</option>
                      </select>
                        <select name='city'  style="width:90px;">
                        </select>
                        <script src="<%=SitePath%>js/getcity.js"></script>
                        <script>initprovcity('','');</script></td>
                  </tr>
                  <!--<tr>
                    <td height="30">QQ���룺</td>
                    <td><input name="UserQQ"  class="borderall" size="25" maxlength="10"  onchange="UserQQ_enter();" onKeyUp="UserQQ_enter();" onBlur="UserQQ_enter();" />
                        <font color="#FF0000">*</font>��<span id="enter_qq"></span></td>
                  </tr>-->
                  <tr align="middle">
                    <td colspan="2" height="30"><script src="<%=SitePath%>inc/ValidateClass.asp?act=showvalidatelaoy" type="text/javascript"></script><input id="Action" type="hidden" value="SaveReg1" name="Action" />
                        <input name="Submit2" type="submit"  class="borderall" value=" ע�� " />
                        <input name="Reset" type="reset"  class="borderall" id="Reset" value=" ������д " />                    </td>
                  </tr>
                </form>
              </table>
		  </div>
		</div>
	</div>
	<div id="nw_right">
		<%Echo ShowAD(3)%>
        <div id="web2r">
			<h5>��������</h5>
			<ul id="list10">
            	<%Call ShowArticle(0,10,5,"��",100,"no","DateAndTime desc,ID desc",0,1,0)%>
            </ul>
  		</div>
	</div>
</div>
<%=Copy%>
</div>
</body>
</html>