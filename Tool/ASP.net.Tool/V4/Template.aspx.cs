using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.ApplicationBlocks.Data;



public partial class _$FileName$ : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            string act = Request["act"];
            string id = Request["id"];

            switch (act)
            {
                case "Add":
                    MultiView1.SetActiveView(ViewEdit);
                    OptAdd();
                    break;
                case "Edit":
                    MultiView1.SetActiveView(ViewEdit);
                    OptEdit(id);
                    break;
                case "Del":
                    OptDel(id);
                    MultiView1.SetActiveView(ViewList);
                    OptList();
                    break;
                case "Show":
                    OptShow(id);
                    MultiView1.SetActiveView(ViewShow);
                    break;
                default:
                    MultiView1.SetActiveView(ViewList);
                    OptList();
                    break;
            }
        }
    }

    protected void OptAdd()
    {
        lblTitle.Text = "添加";
        btnSave.Text = "添加";
        lblDaoHang.Text = "  :: 添加";
        //$AddCode$
    }

    protected void OptEdit(string id)
    {
        lblTitle.Text = "修改";
        btnSave.Text = "修改";
        lblDaoHang.Text = "  :: 修改";
        DataRowView drv = SqlHelper.ExecuteDataset(SqlHelper.SqlConnection, CommandType.Text, "select * from [$TableName$] where Id=" + id).Tables[0].DefaultView[0];
        //$EditCode$
        txtIDEdit.Value = drv["Id"].ToString();
        txtFromUrl.Value = Request.UrlReferrer.OriginalString;
    }

    protected void OptShow(string id)
    {

        DataRowView drv = SqlHelper.ExecuteDataset(SqlHelper.SqlConnection, CommandType.Text, "select * from [$TableName$] where Id=" + id).Tables[0].DefaultView[0];
        //$ShowCode$
    }

    protected void OptDel(string id)
    {

        SqlHelper.ExecuteNonQuery(SqlHelper.SqlConnection, CommandType.Text, "delete from [$TableName$] where Id=" + id);
        Response.Redirect(Request.UrlReferrer.OriginalString, true);
    }

    protected void OptList()
    {
        //$ListCode$

    }

    protected void btnSave_Click(object sender, EventArgs e)
    {
        //$btnSave$
        if (txtIDEdit.Value == string.Empty)
        {
            SqlHelper.ExecuteNonQuery(SqlHelper.SqlConnection, CommandType.Text, "insert into [$TableName$]($Fields$) values ($FieldsValue$)", parms);
            Response.Redirect(Request.Path.ToString());
        }
        else
        {
            parms[parms.Length - 1] = new SqlParameter("@Id", txtIDEdit.Value.Trim());
            SqlHelper.ExecuteNonQuery(SqlHelper.SqlConnection, CommandType.Text, "update [$TableName$] set $UpdateFields$ where Id=@Id", parms);
            Response.Redirect(Request["txtFromUrl"]);
        }
    }

    protected void btnSearch_Click(object sender, EventArgs e)
    {
        Response.Redirect(Request.Path 
        	+ "?txtTitle=" //+ Server.UrlEncode(txtTitle.Text)
			+ "&txtT1=" + Server.UrlEncode(Request.Form["txtT1"])
			+ "&txtT2=" + Server.UrlEncode(Request.Form["txtT2"])
//$SearchCodeURL$


            
            , true);
    }



    protected void AspNetPager1_PageChanged(object sender, EventArgs e)
    {

        string txtSQL = @"
a.* from [$TableName$] a  where 1=1 

";
//$SearchCode$

        txtT1.Text = Request["txtT1"];
        if (!string.IsNullOrEmpty(txtT1.Text.Trim()))
        {
            txtSQL += " and AddTime>='" + txtT1.Text.Trim() + " 00:00:00'";
        }

        txtT2.Text = Request["txtT2"];
        if (!string.IsNullOrEmpty(txtT2.Text.Trim()))
        {
            txtSQL += " and AddTime<='" + txtT2.Text.Trim() + " 23:59:59'";
        }
        
        AspNetPager1.PageSize = 30;

        SqlParameter[] parms ={
            new SqlParameter("@strSQL", txtSQL),
            new SqlParameter("@strOrder", " a.id desc"),
            new SqlParameter("@PageSize", AspNetPager1.PageSize),
            new SqlParameter("@PageIndex", Request["Page"]==null?"1":Request["Page"])
            };


        DataSet ds = SqlHelper.ExecuteDataset(SqlHelper.SqlConnection, CommandType.StoredProcedure, "ShowPage", parms);

        AspNetPager1.RecordCount = Convert.ToInt32(ds.Tables[1].Rows[0][0]);

        GridView1.DataSource = ds.Tables[0];
        GridView1.DataBind();

    }
}
