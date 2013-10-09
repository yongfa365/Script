<%@ Page Language="C#" AutoEventWireup="true" CodeFile="$FileName$.aspx.cs" EnableEventValidation="false" Inherits="_$FileName$" %>
<%@ Register Assembly="AspNetPager" Namespace="Wuqi.Webdiyer" TagPrefix="webdiyer" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="pragma" content="no-cache" /> 
<link href="/Style.css" rel="stylesheet" type="text/css" /> 
<script language="javascript" type="text/javascript" src="/Images/Comm.js"></script> 
<script language="javascript" type="text/javascript" src="/My97DatePicker/WdatePicker.js" ></script>
<script language="javascript" type="text/javascript">
function del(id) {
    if (confirm("确定删除？")) {
        window.location = "?act=Del&id=" + id + "&ReturnUrl=" + encodeURI(window.location.href);
    }
}

</script>
<title></title>
</head>
<body>
    <form id="yongfa365_jhsb" runat="server">
    <div id="daohang"><asp:Literal ID="lblDaoHang" runat="server"></asp:Literal></div>
    <input type="button" value="<< 返回" onclick="window.history.go(-1)" />　　　　
    <input type="button" value="添加" onclick="window.location = '?act=Add'" />
    <asp:MultiView ID="MultiView1" runat="server">
        <asp:View ID="ViewEdit" runat="server">
<script language="javascript" type="text/javascript">
<!--
function FormValidator(theForm)
{ 
$ValidatorHtml$
return true;
}
-->
</script>
        <table>
		  <tr>
		    <th colspan="2"><asp:Label ID="lblTitle" runat="server" ></asp:Label></th>
		  </tr>
$EditHtml$
		  <tr>
		    <td>&nbsp;</td>
		    <td>
		    <asp:Button ID="btnSave" runat="server" onclick="btnSave_Click" OnClientClick="return FormValidator(yongfa365_jhsb)" />
		    <asp:HiddenField ID="txtFromUrl" runat="server" ></asp:HiddenField>
		    <asp:HiddenField ID="txtIDEdit" runat="server"></asp:HiddenField>
		    </td>
		  </tr>
		</table>
        </asp:View>

        <asp:View ID="ViewShow" runat="server">
        <table>
		  <tr>
		    <th colspan="2">详细信息</th>
		  </tr>
$ShowHtml$
		</table>
        </asp:View>

        <asp:View ID="ViewList" runat="server">
         <table>
		  <tr>
		    <th>查询</th>
		  </tr>
		  <tr>
		    <td>
		    $SearchHtml$
		    录入时间：从
      <asp:TextBox ID="txtT1" runat="server" onFocus="WdatePicker()" Width="75px"></asp:TextBox>
      到
      <asp:TextBox ID="txtT2" runat="server" onFocus="WdatePicker()" Width="75px"></asp:TextBox>
		    <asp:Button ID="btnSearch" runat="server" OnClick="btnSearch_Click" Text="查询" />
		    </td>
		  </tr>
		</table>
<asp:Repeater ID="GridView1" runat="server" EnableViewState="False" >
    <HeaderTemplate>
        <table cellspacing="0" rules="all" border="1" id="GridView1" style="border-collapse:collapse;">
         <tr>
<th>编号</th>
$HeaderTextHtml$
<th>编辑</th>
<th>删除</th>
            </tr>
    </HeaderTemplate>
    <ItemTemplate>
        <tr>
<td><a href="?act=Show&amp;id=<%#Eval("Id") %>"><%#Eval("Id") %></a></td>
$ListHtml$
<td><a href="?act=Edit&amp;id=<%#Eval("Id") %>">编辑</a></td>
<td><a href="###" onclick="del(<%# Eval("ID") %>)">删除</a></td>
        </tr>
    </ItemTemplate>
    <FooterTemplate>
        </table></FooterTemplate>
</asp:Repeater>

        

        <webdiyer:AspNetPager ID="AspNetPager1" ShowPageIndexBox="Never" runat="server" UrlPaging="True"
            FirstPageText="首页" LastPageText="尾页" NextPageText="下一页" PrevPageText="上一页" OnPageChanged="AspNetPager1_PageChanged"
            CustomInfoHTML="共%RecordCount%条信息 每页%PageSize%条　当前%CurrentPageIndex%/%PageCount%页"
            ShowCustomInfoSection="Right" CurrentPageButtonClass="current" CssClass="AspNetPager1"
            CustomInfoSectionWidth="">
        </webdiyer:AspNetPager>
       </asp:View>
    </asp:MultiView>
    </form>
</body>
</html>


