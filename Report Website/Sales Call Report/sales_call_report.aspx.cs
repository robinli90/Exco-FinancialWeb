using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.Odbc;
using ExcoUtility;

public partial class sales_call_report : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
    }

    // add a sales call report
    protected void AddButton_Click(object sender, EventArgs e)
    {
        Session["_SalesCallReportID"] = "0";
        Response.Redirect("submit_sales_call_report.aspx");
    }

    // check all legacy reports
    protected void CheckButton_Click(object sender, EventArgs e)
    {
        //Response.Write("<SCRIPT LANGUAGE=\"JavaScript\">alert(\"Submit report failed! Window's account: " + Environment.UserName + "\")</SCRIPT>");
        // create style
        GridViewReports.Visible = true;
        GridViewReports.Columns.Clear();
        GridViewReports.CaptionAlign = TableCaptionAlign.Top;
        GridViewReports.Caption = "<font size=\"5\" color=\"red\">Legacy Sales Call Report for " + CustomerList.Text + " </font>";
        // create grid view table
        DataTable report = new DataTable();
        report.Columns.Add("ID");
        report.Columns.Add("Visit Date");
        report.Columns.Add("Follow Up Date");
        report.Columns.Add("Sales");
        string query = "select [ID], [Date], [FollowUpDate], [Sales] from [tiger].[dbo].[Sales_Call_Report] where [CustomerName]='" + CustomerList.Text.Replace("'", "''") + "'";
        ExcoODBC database = ExcoODBC.Instance;
        database.Open(Database.DECADE_MARKHAM);
        OdbcDataReader reader = database.RunQuery(query);
        while (reader.Read())
        {
            DataRow row = report.NewRow();
            row[0] = reader[0].ToString();
            row[1] = reader[1].ToString();
            row[2] = reader[2].ToString();
            row[3] = reader[3].ToString();
            report.Rows.Add(row);
        }
        reader.Close();
        // bind data
        report.AcceptChanges();
        GridViewReports.AutoGenerateSelectButton = true;
        GridViewReports.DataSource = report;
        GridViewReports.DataBind();
    }

    // display the report detail
    protected void Select_Click(object sender, EventArgs e)
    {
        Session["_SalesCallReportID"] = GridViewReports.SelectedRow.Cells[1].Text;
        Response.Redirect("submit_sales_call_report.aspx");
    }

}