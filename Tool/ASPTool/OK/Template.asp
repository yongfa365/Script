<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="Admin.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript" src="Admin.js"></script>
<title>柳永法 ASP生成工具</title>
</head>
<body>
<!--#include file="conn.asp" -->
<!--#include file="Page.asp" -->
<%
TableName = "$TableName$"
ID = SafeRequest("ID")
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
    <td>添加</td>
  </tr>
  <tr>
    <td><input type="button"  value="&lt;-返回-&gt;" onclick="window.history.go(-1)" />
      <input type="button" value="列表首页" onclick="window.location.href='<%=NowFile%>'" />
    </td>
  </tr>
</table>
<table cellpadding="0" cellspacing="0" align="center" class="Admin_Table">
  <form action="<%=NowFile%>?action=addok" method="post" name="form1" id="form1" >
    $Admin_ADDHTML$
    <tr>
      <td colspan="2" align="center"><input type="submit" name="Submit3" value="添加" />
        <input type="reset" name="Submit4" value="重置" /></td>
    </tr>
  </form>
</table>
<%End Sub%>
<%Sub modi()%>
<%
Set rs = Server.CreateObject("Adodb.Recordset")
sql = "select * from [" & TableName & "] where  id=" & request("id")
rs.Open sql, conn, 1, 3
If Not rs.EOF Then
%>
<table cellpadding="0" cellspacing="0" align="center" class="Admin_Table">
  <tr class="Admin_Title">
    <td>修改</td>
  </tr>
  <tr>
    <td><input type="button"  value="&lt;-返回-&gt;" onclick="window.history.go(-1)" />
      <input type="button" value="列表首页" onclick="window.location.href='<%=NowFile%>'" />
    </td>
  </tr>
</table>
<table cellpadding="0" cellspacing="0" align="center" class="Admin_Table">
  <form action="<%=NowFile%>?action=modiok" method="post" name="form1" id="form1" >
    $Admin_EditHTML$
    <tr>
      <td colspan="2" align="center"><input type="hidden" name="ID" value="<%=rs("ID")%>" />
        <input type="hidden" name="ModiURL" value="<%=ModiURL%>" />
        <input type="submit" name="Submit" value="修改" />
        &nbsp;&nbsp;
        <input type="reset" name="Submit2" value="重置" />
      </td>
    </tr>
  </form>
</table>
<%end if%>
<%End Sub%>
<%Sub show()%>
<%
Set rs = Server.CreateObject("Adodb.Recordset")
sql = "select * from [" & TableName & "] where  id=" & request("id")
rs.Open sql, conn, 1, 1
If Not rs.EOF Then
%>
<table cellpadding="0" cellspacing="0" align="center" class="Admin_Table">
  <tr class="Admin_Title">
    <td>详细信息</td>
  </tr>
  <tr>
    <td><input type="button"  value="&lt;-返回-&gt;" onclick="window.history.go(-1)" />
      <input type="button" value="列表首页" onclick="window.location.href='<%=NowFile%>'" />
    </td>
  </tr>
</table>
<table cellpadding="0" cellspacing="0" align="center" class="Admin_Table">
  $Admin_ShowHTML$
  <tr>
    <td colspan="2" align="center"><input type="button"  value="修改" onclick="window.location.href='<%=NowFile%>?action=modi&id=<%=id%>&ModiURL=<%=server.urlencode(ModiURL)%>'" />
      <input type="button"  value="&lt;-返回-&gt;" onclick="window.history.go(-1)" />
    </td>
  </tr>
</table>
<%end if%>
<%End Sub%>
<%Sub list()%>
<%
'搜索条件
SearchURL=""
sql = ""
$Admin_SearchHTML_ASP$

'response.write sql & "<br>"
'response.write SearchURL & "<br>"
%>
<table cellpadding="0" cellspacing="0" align="center" class="Admin_Table">
  <tr class="Admin_Title">
    <td>列表</td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <form id="form1" name="form1" method="get" action="<%=NowFile%>?action=list<%=SearchURL%>">
          <tr>
            <td><input type="button" value="添加" onclick="window.location.href='?action=add'" />
            <input type="button" value="列表首页" onclick="window.location.href='<%=NowFile%>'" />
              $Admin_SearchHTML$
              <input type="submit"  value="查找" />
              &nbsp;</td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>

<table cellpadding="0" cellspacing="0"  align="center"  class="Admin_Table" id="trColor">
  <tr>
    <th>编号</th>
    <th>点击</th>
    $Admin_ListHTML1$
    <th>发布时间</th>
    <th>操作<input type="checkbox" onclick="CheckBox('trColor')" title="反选"><a href="javascript:DelChkID('<%=NowFile%>?action=delok&ID=','trColor')" title="删除所有选中记录">删</a></th>
  </tr>
<%
'===========================分页设置===========================
Set Page1 = New Page
Page1.CurrentPage = Trim(Request("Page"))
Page1.PageNum = 20
Page1.IsDeBug = True
Page1.Url = NowFile & "?dev=www.yongfa365.com" & SearchURL & "&"
Page1.Exec TableName, "id", 1, "" , " 1=1 " & sql, conn

Set rs = Page1.PageRs
'===========================分页设置===========================
For i = 1 To Page1.PageNum
If rs.EOF Then Exit For
%>
  <tr>
    <td><a href="<%=NowFile%>?action=show&id=<%=rs("id")%>"><%=rs("id")%></a></td>
    <td><%=rs("hits")%></td>
    $Admin_ListHTML2$
    <td><%=rs("addtime")%></td>
    <td><a href="<%=NowFile%>?action=modi&ID=<%=rs("ID")%>">修改</a> <a href="<%=NowFile%>?action=delok&ID=<%=rs("ID")%>" onclick="return confirm('确定删除？')">删除</a> <input type="checkbox" value="<%=rs("ID")%>"></td>
  </tr>
<%
rs.movenext
Next
%>
</table>
<%
response.flush()
'===========================分页显示===========================
Page1.Temp="{N1} {N2} {N3} {N4} || 共:{N8}条记录 {N6}页 当前为第{N5}页 每页{N7}条"
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
Page1.Temp="{N1}{N2}{N9}{L}{N}{L/}{N10}{N3}{N4} 共{N8}条记录 {N6}页 当前为第{N5}页 每页{N7}条"
Page1.show(4)
response.write("<p><p><p>")
'===========================分页显示===========================
%>
<%end sub %>
<%sub addok()%>
<%

Set rs = Server.CreateObject("Adodb.Recordset")
sql = "select  * from [" & TableName & "] where 1=2"
rs.Open sql, conn, 1, 3
rs.addnew()
$InDB_VBS$
rs("Hits") = 0
rs("ADDTime") = Now()
rs("ModiTime") = Now()
rs("OrderID") = 0
rs("Flags") = 0
rs("Remark") = ""
rs.update
rs.Close
Set rs = Nothing
%>
<%end sub%>
<%sub modiok()%>
<%

Set rs = Server.CreateObject("Adodb.Recordset")
sql = "select  * from [" & TableName & "] where ID="&ID
rs.Open sql, conn, 1, 3
If Not rs.EOF Then
    $InDB_VBS$
    '	rs("Hits")=0
    '	rs("ADDTime")=now()
    rs("ModiTime") = Now()
    '	rs("OrderID")=0
    '	rs("Flags")=0
    '	rs("Remark")=""
    rs.update
End If
rs.Close
Set rs = Nothing
%>
<%end sub%>
<%sub delok()%>
<%
conn.Execute ("delete from [" & TableName & "] where ID in(" & ID & ")")
%>
<%end sub%>
<script type="text/javascript">window.onload = function(){GridViewColor("trColor","#fff","#eee","#6df","#fd6");}</script>
</body>
</html>
