using System.Drawing;
using System;
using System.IO;
using System.Data;
using System.Data.Odbc;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ExcoUtility;

public partial class invoice_register : System.Web.UI.Page
{

    //private string fiscalYear = "ALL" , plant = "Markham" , size = "18"; Test
    private string fiscalYear, plant, size;
    private string fromYear, toYear;
    private string[,] die_info = new string[150,3];
    private string[,] die_info_all = new string[1000,4];
    
    // Return the appropriate query statement
    public string Query(string plant, string size, string startYear, string endYear, string custname)
    {
        int size_comp = Convert.ToInt32(size);
        double size_upper = Convert.ToDouble(size_comp) + 0.2;
        double size_lower = size_upper - 1;
        string start_date = startYear + "-10-01";
        string end_date = endYear + "-09-30";

        string custstr = "";
        if (custname != "")
        {
            custstr = " and c.name like '%" + custname + "%' ";
        }
        string tempsql = "select b.customercode + ' - - ' + c.name as Customer, " +
                "count(b.customercode + ' - - ' + c.name) as Count, " +
                "avg(a.qty) as QTY, sum(total) as Total " +
                "from d_orderitem as a " +
                "inner join d_order as b on a.ordernumber=b.ordernumber " +
                "inner join d_customer as c on b.customercode=c.customercode where a.suffix like '%0X00%'" + custstr +
                "and (cast(substring(a.suffix,3,7) as decimal(8,3)) " +
                "between " + size_lower.ToString() + " and " + size_upper.ToString() + " " +
                "or  (cast(substring(a.suffix,3,7) as decimal(8,3))/25.4) " +
                "between " + size_lower.ToString() + " and " + size_upper.ToString() + ") " +
                "and invoicedate between '" + start_date + "' and '" + end_date + "' " +
                "group by (b.customercode + ' - - ' + c.name) order by Count desc";

        return tempsql;
    }

    // Return the query for the "ALL" selection
    public string All_Query(string plant, string size, string startYear, string endYear, string custname)
    {
        string start_date = startYear + "-10-01";
        string end_date = endYear + "-09-30";
        string custstr = "";
        if (custname != "")
        {
            custstr = " and c.name like '%" + custname + "%' ";
        }
        string tempsql = "select a.Customer as Cust, a.size as Size, count(a.Customer) as Count, sum(a.Total) as Total from ( " +
                "select b.customercode + ' - - ' + c.name as Customer, total as Total, " +
		        "case when cast(substring(substring(a.suffix,3,7),0,charindex('.', substring(a.suffix,3,7))) as int) > 50 then " +
                "cast(round(cast(substring(substring(a.suffix,3,7),0,charindex('.', substring(a.suffix,3,7))) as int)/25.4,0) as int) " +
                "else cast(round(cast(substring(substring(a.suffix,3,7),0,charindex('.', substring(a.suffix,3,7))) as int),0) as int) " +
                "end as Size from d_orderitem as a " +
		        "inner join d_order as b on a.ordernumber=b.ordernumber " +
		        "inner join d_customer as c on b.customercode=c.customercode where a.suffix like '%0X00%' " + custstr +
		        "and invoicedate between '" + start_date + "' and '" + end_date + "' ) a " + 
                "group by a.Customer, a.size order by Cust, Size ";

        

        return tempsql;
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        Session["_SalesCallReportID"] = 0;
    }

    protected void ButtonGenerate_Click(object sender, EventArgs e)
    {
        LblNofound.Visible = false;
        string customername = "";
        if (ChkCust.Checked)
        {
            customername = CustomerList.Text;
            LblNofound.Visible = true;
        }

        ReportGrid.Visible = true;
        DataTable report = new DataTable();

        plant = PlantList1.SelectedValue.ToString();
        size = SizeList1.SelectedValue.ToString();
        fiscalYear = YearChoice1.SelectedValue.ToString();

        if (fiscalYear == "All" || fiscalYear == "ALL")
        {
            fromYear = "2010";
            toYear = "2017";
        }
        else
        {
            fromYear = (Convert.ToInt32(fiscalYear) - 1).ToString();
            toYear = fiscalYear;
        }

        ExcoODBC database = ExcoODBC.Instance;
        OdbcDataReader reader;
        if (plant == "Markham")
        {
            database.Open(Database.DECADE_MARKHAM);
        }
        else if (plant == "Michigan")
        {
            database.Open(Database.DECADE_MICHIGAN);
        }
        else if (plant == "Texas")
        {
            database.Open(Database.DECADE_TEXAS);
        }
        else if (plant == "Colombia")
        {
            database.Open(Database.DECADE_COLOMBIA);
        }

        // If size is selective, then execute the following solo query
        if (!(size == "ALL" || size == "All"))
        {


            int iterator = 0;
            reader = database.RunQuery(Query(plant, size, fromYear, toYear, customername));
            while (reader.Read())
            {
                die_info[iterator, 0] = reader[0].ToString();
                die_info[iterator, 1] = reader[1].ToString();
                die_info[iterator, 2] = reader[3].ToString();
                iterator++;
            }
            reader.Close();
            if (Convert.ToInt32(toYear) - Convert.ToInt32(fromYear) > 2) toYear = "all years";
            GridView gridView = ReportGrid;
            gridView.Caption = "<font size=\"5\" color=\"red\">" + size + "\" purchases in " + toYear.ToString() + " (" + plant + ")</font>";
            report.Columns.Add("Customer");
            report.Columns.Add("Number of dies purchased");
            report.Columns.Add("Total");
            List<DataRow> rowList = new List<DataRow>();
            for (int i = 0; i < iterator; i++)
            {
                DataRow row = report.NewRow();
                row["Customer"] = die_info[i, 0];
                row["Number of dies purchased"] = die_info[i, 1];
                row["Total"] = die_info[i, 2];
                report.Rows.Add(row);
            }

            if (iterator == 0)
            {
                LblNofound.Text = "<font size=\"5\" color=\"red\">No Result Found for " + customername + "</font>";
            }

            gridView.DataSource = report;
            gridView.DataBind();

        }
        else // All size query
        {
            int iterator = 0;
            reader = database.RunQuery(All_Query(plant, size, fromYear, toYear, customername));
            while (reader.Read())
            {
                die_info_all[iterator, 0] = reader[0].ToString();
                die_info_all[iterator, 1] = reader[1].ToString();
                die_info_all[iterator, 2] = reader[2].ToString();
                die_info_all[iterator, 3] = reader[3].ToString();
                iterator++;
            }
            reader.Close();
            if (Convert.ToInt32(toYear) - Convert.ToInt32(fromYear) > 2) toYear = "all years";
            GridView gridView = ReportGrid;
            gridView.Caption = "<font size=\"5\" color=\"red\">Die purchases by customer in " + toYear.ToString() + " (" + plant + ")</font>";
            report.Columns.Add("Customer");
            report.Columns.Add("Size");
            report.Columns.Add("Number of dies purchased");
            report.Columns.Add("Total");
            List<DataRow> rowList = new List<DataRow>();
            string prev_cust = "";
            for (int i = 0; i < iterator; i++)
            {
                DataRow row = report.NewRow();
                if (!(die_info_all[i,0] == prev_cust)) {
                    row["Customer"] = die_info_all[i, 0];
                    prev_cust = die_info_all[i, 0];
                } else {
                    row["Customer"] = "  ";
                }
                row["Size"] = die_info_all[i, 1];
                row["Number of dies purchased"] = die_info_all[i, 2];
                row["Total"] = die_info_all[i, 3];
                report.Rows.Add(row);
            }

            if (iterator == 0)
            {
                LblNofound.Text = "<font size=\"5\" color=\"red\">No Result Found for " + customername + "</font>";
            }
            gridView.DataSource = report;
            gridView.DataBind();

        }
    }

    private bool GenerateCSVFile(ref StreamWriter writer, GridView gridView)
    {
        bool hasContent = false;
        string title = gridView.Caption;
        title = title.Replace("<font size=\"9\" color=\"red\">", "");
        title = title.Replace("</font>", "");
        writer.WriteLine(title);
        // write header
        foreach (TableCell cell in gridView.HeaderRow.Cells)
        {
            writer.Write(cell.Text + ",");
        }
        writer.WriteLine();
        // write content
        foreach (GridViewRow row in gridView.Rows)
        {
            string content = string.Empty;
            foreach (TableCell cell in row.Cells)
            {
                content += "\"" + cell.Text + "\",";
                hasContent = true;
            }
            content = content.Replace("&#39;", "'");
            content = content.Replace("&amp;", "&");
            content = content.Replace("&nbsp;", " ");
            writer.WriteLine(content);
        }
        writer.WriteLine();
        writer.WriteLine();
        writer.WriteLine();
        return hasContent;
    }

    protected void ButtonSave_Click(object sender, EventArgs e)
    {
        // write to excel
        string fileName = @"C:\Invoice Register\" + User.Identity.Name + " Customer Die Report for " + PlantList1.SelectedValue.ToString() + " at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + ".csv";

        if (Directory.Exists(@"C:\Invoice Register\"))
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
        }
        else
        {
            Directory.CreateDirectory(@"C:\Income Statement Report\");
        }

        StreamWriter writer = new StreamWriter(fileName);
        if (ReportGrid.Visible)
        {
            GenerateCSVFile(ref writer, ReportGrid);
        }
        writer.Flush();
        writer.Close();
        // provide file to download
        HttpResponse response = HttpContext.Current.Response;
        response.AppendHeader("Content-Disposition", "attachment; filename=" + fileName);
        FileInfo file = new FileInfo(fileName);
        response.AppendHeader("Content-Length", file.Length.ToString());
        response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        response.WriteFile(fileName);
        response.End();
    
    }
}