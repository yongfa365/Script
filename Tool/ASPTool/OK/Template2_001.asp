<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="Admin.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript" src="Admin.js"></script>
<title>������ ASP���ɹ���</title>
</head>
<body>
<!--#include file="conn.asp" -->
<!--#include file="Page.asp" -->
<%
TableName = "$TableName$"
$AutoFieldName$ = SafeRequest("$AutoFieldName$")
NowFile = request.ServerVariables("script_name")
ModiURL = request("ModiURL")
If ModiURL = "" Then ModiURL = Request.ServerVariables("HTTP_REFERER")
action = request("action")
Select Case action
    Case "add"
        Call Add()
    Case "addok"
        Call addok()
        response.redirect NowFile
    Case "modi"
        Call modi()
    Case "modiok"
        Call modiok()
        response.redirect ModiURL
    Case "delok"
        Call delok()
        response.redirect ModiURL
    Case "show"
        Call show()
    Case Else
        Call list()
End Select

%>
<%Sub add()%>
<table cellpadding="0" cellspacing="0" align="center" class="Admin_Table">
  <tr class="Admin_Title">
    <td>���</td>
  </tr>
  <tr>
    <td><input type="button"  value="&lt;-����-&gt;" onclick="window.history.go(-1)" />
      <input type="button" value="�б���ҳ" onclick="window.location.href='<%=NowFile%>'" />
    </td>
  </tr>
</table>
<table cellpadding="0" cellspacing="0" align="center" class="Admin_Table">
  <form action="<%=NowFile%>?action=addok" method="post" name="form1" id="form1" >
    $Admin_ADDHTML$
    <tr>
      <td colspan="2" align="center"><input type="submit" name="Submit3" value="���" />
        <input type="reset" name="Submit4" value="����" /></td>
    </tr>
  </form>
</table>
<%End Sub%>
<%Sub modi()%>
<%
Set rs = Server.CreateObject("Adodb.Recordset")
sql = "select * from [" & TableName & "] where  $AutoFieldName$=" & request("$AutoFieldName$")
rs.Open sql, conn, 1, 3
If Not rs.EOF Then
%>
<table cellpadding="0" cellspacing="0" align="center" class="Admin_Table">
  <tr class="Admin_Title">
    <td>�޸�</td>
  </tr>
  <tr>
    <td><input type="button"  value="&lt;-����-&gt;" onclick="window.history.go(-1)" />
      <input type="button" value="�б���ҳ" onclick="window.location.href='<%=NowFile%>'" />
    </td>
  </tr>
</table>
<table cellpadding="0" cellspacing="0" align="center" class="Admin_Table">
  <form action="<%=NowFile%>?action=modiok" method="post" name="form1" id="form1" >
    $Admin_EditHTML$
    <tr>
      <td colspan="2" align="center"><input type="hidden" name="$AutoFieldName$" value="<%=rs("$AutoFieldName$")%>" />
        <input type="hidden" name="ModiURL" value="<%=ModiURL%>" />
        <input type="submit" name="Submit" value="�޸�" />
        &nbsp;&nbsp;
        <input type="reset" name="Submit2" value="����" />
      </td>
    </tr>
  </form>
</table>
<%end if%>
<%End Sub%>
<%Sub show()%>
<%
Set rs = Server.CreateObject("Adodb.Recordset")
sql = "select * from [" & TableName & "] where  $AutoFieldName$=" & request("$AutoFieldName$")
rs.Open sql, conn, 1, 1
If Not rs.EOF Then
%>
<table cellpadding="0" cellspacing="0" align="center" class="Admin_Table">
  <tr class="Admin_Title">
    <td>��ϸ��Ϣ</td>
  </tr>
  <tr>
    <td><input type="button"  value="&lt;-����-&gt;" onclick="window.history.go(-1)" />
      <input type="button" value="�б���ҳ" onclick="window.location.href='<%=NowFile%>'" />
    </td>
  </tr>
</table>
<table cellpadding="0" cellspacing="0" align="center" class="Admin_Table">
  $Admin_ShowHTML$
  <tr>
    <td colspan="2" align="center"><input type="button"  value="�޸�" onclick="window.location.href='<%=NowFile%>?action=modi&$AutoFieldName$=<%=$AutoFieldName$%>&ModiURL=<%=ModiURL%>'" />
      <input type="button"  value="&lt;-����-&gt;" onclick="window.history.go(-1)" />
    </td>
  </tr>
</table>
<%end if%>
<%End Sub%>
<%Sub list()%>
<%
'��������
SearchURL=""
sql = ""
$Admin_SearchHTML_ASP$

'response.write sql & "<br>"
'response.write SearchURL & "<br>"
%>
<table cellpadding="0" cellspacing="0" align="center" class="Admin_Table">
  <tr class="Admin_Title">
    <td>�б�</td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <form id="form1" name="form1" method="get" action="<%=NowFile%>?action=list<%=SearchURL%>">
          <tr>
            <td><input type="button" value="���" onclick="window.location.href='?action=add'" />
            <input type="button" value="�б���ҳ" onclick="window.location.href='<%=NowFile%>'" />
              $Admin_SearchHTML$
              <input type="submit"  value="����" />
              &nbsp;</td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>

<table cellpadding="0" cellspacing="0"  align="center"  class="Admin_Table" id="trColor">
  <tr>
    <th>���</th>
    $Admin_ListHTML1$
    <th>����<input type="checkbox" onclick="CheckBox('trColor')" title="��ѡ"><a href="javascript:DelChkID('<%=NowFile%>?action=delok&$AutoFieldName$=','trColor')" title="ɾ������ѡ�м�¼">ɾ</a></th>
  </tr>
<%
'===========================��ҳ����===========================
Set Page1 = New Page
Page1.CurrentPage = Trim(Request("Page"))
Page1.PageNum = 20
Page1.IsDeBug = True
Page1.Url = NowFile & "?dev=www.yongfa365.com" & SearchURL & "&"
Page1.Exec TableName, "$AutoFieldName$", 1, "" , " 1=1 " & sql, conn

Set rs = Page1.PageRs
'===========================��ҳ����===========================
For i = 1 To Page1.PageNum
If rs.EOF Then Exit For
%>
  <tr>
    <td><a href="<%=NowFile%>?action=show&$AutoFieldName$=<%=rs("$AutoFieldName$")%>"><%=rs("$AutoFieldName$")%></a></td>
    $Admin_ListHTML2$
    <td><a href="<%=NowFile%>?action=modi&$AutoFieldName$=<%=rs("$AutoFieldName$")%>">�޸�</a> <a href="<%=NowFile%>?action=delok&$AutoFieldName$=<%=rs("$AutoFieldName$")%>" onclick="return confirm('ȷ��ɾ����')">ɾ��</a> <input type="checkbox" value="<%=rs("$AutoFieldName$")%>"></td>
  </tr>
<%
rs.movenext
Next
%>
</table>
<%
response.flush()
'===========================��ҳ��ʾ===========================
Page1.Temp="{N1} {N2} {N3} {N4} || ��:{N8}����¼ {N6}ҳ ��ǰΪ��{N5}ҳ ÿҳ{N7}��"
Page1.show(1)
response.write("<p><p><p>")
Page1.Temp="[{N1}] [{N2}] [{N3}] [{N4}]"
Page1.show(1)
response.write("<p><p><p>")
Page1.Temp="{N1} {N2} {L}{N}{L/}{N3}{N4}"
Page1.show(2)
response.write("<p><p><p>")
Page1.Temp="{N1}{N2}{N9}{L}{N}{L/}{N10}{N3}{N4}"
Page1.show(3)
response.write("<p><p><p>")
Page1.Temp="{N1}{N2}{N9}{L}{N}{L/}{N10}{N3}{N4} ��{N8}����¼ {N6}ҳ ��ǰΪ��{N5}ҳ ÿҳ{N7}��"
Page1.show(4)
response.write("<p><p><p>")
'===========================��ҳ��ʾ===========================
%>
<%end sub %>
<%sub addok()%>
<%

Set rs = Server.CreateObject("Adodb.Recordset")
sql = "select  * from [" & TableName & "] where 1=2"
rs.Open sql, conn, 1, 3
rs.addnew()
$InDB_VBS$
rs.update
rs.Close
Set rs = Nothing
%>
<%end sub%>
<%sub modiok()%>
<%

Set rs = Server.CreateObject("Adodb.Recordset")
sql = "select  * from [" & TableName & "] where $AutoFieldName$="&$AutoFieldName$
rs.Open sql, conn, 1, 3
If Not rs.EOF Then
    $InDB_VBS$
    rs.update
End If
rs.Close
Set rs = Nothing
%>
<%end sub%>
<%sub delok()%>
<%
conn.Execute ("delete from [" & TableName & "] where $AutoFieldName$ in(" & $AutoFieldName$ & ")")
%>
<%end sub%>
<script type="text/javascript">window.onload = function(){GridViewColor("trColor","#fff","#eee","#6df","#fd6");}</script>
</body>
</html>
