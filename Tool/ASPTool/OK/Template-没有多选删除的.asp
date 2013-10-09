<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="Admin.css" rel="stylesheet" type="text/css" />
<title>ASP生成工具</title>
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
    <th>操作</th>
  </tr>
<%
'===========================分页设置===========================
Set Page1 = New Page
Page1.CurrentPage = Trim(Request("Page"))
Page1.PageNum = 20
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
    <td><a href="<%=NowFile%>?action=modi&ID=<%=rs("ID")%>">修改</a> <a href="<%=NowFile%>?action=delok&ID=<%=rs("ID")%>" onclick="return confirm('确定删除？')">删除</a> </td>
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
<script type="text/javascript">
    //把事件放在onload里，因为我不知道JS如果直接写到这儿是不是会等页面加载完才执行
    //使用<%=vd%>方式输出GridView的ID是因为某些情况下（如使用了MasterPage）会造成HTML中ID的变化
    //颜色值推荐使用Hex，如 #f00 或 #ff0000
    window.onload = function(){
        GridViewColor("trColor","#fff","#eee","#6df","#fd6");
    }
    
    //参数依次为（后两个如果指定为空值，则不会发生相应的事件）：
    //GridView ID, 正常行背景色,交替行背景色,鼠标指向行背景色,鼠标点击后背景色
    function GridViewColor(GridViewId, NormalColor, AlterColor, HoverColor, SelectColor){
        //获取所有要控制的行
        if(!document.getElementById(GridViewId)){return;}
        var AllRows = document.getElementById(GridViewId).getElementsByTagName("tr");
        
        //设置每一行的背景色和事件，循环从1开始而非0，可以避开表头那一行
        for(var i=1; i<AllRows.length; i++){
            //设定本行默认的背景色
            AllRows[i].style.background = i%2==0?NormalColor:AlterColor;
            
            //如果指定了鼠标指向的背景色，则添加onmouseover/onmouseout事件
            //处于选中状态的行发生这两个事件时不改变颜色
            if(HoverColor != ""){
                AllRows[i].onmouseover = function(){if(!this.selected)this.style.background = HoverColor;}
                if(i%2 == 0){
                    AllRows[i].onmouseout = function(){if(!this.selected)this.style.background = NormalColor;}
                }
                else{
                    AllRows[i].onmouseout = function(){if(!this.selected)this.style.background = AlterColor;}
                }
            }

            //如果指定了鼠标点击的背景色，则添加onclick事件
            //在事件响应中修改被点击行的选中状态
            if(SelectColor != ""){
                AllRows[i].onclick = function(){
                    this.style.background = this.style.background==SelectColor?HoverColor:SelectColor;
                    this.selected = !this.selected;
                }
            }
        }
    }
</script>
</body>
</html>
