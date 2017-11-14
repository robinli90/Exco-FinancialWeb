using System.Drawing;
using System;
using System.Data;
using System.Data.Odbc;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Runtime.InteropServices;
using ExcoUtility;

public partial class Monthly_Sales_By_Customer : System.Web.UI.Page
{
    // customer name
    static string sCustomerName = string.Empty;
    // invoice
    struct Invoice
    {
        public string sInvNum;
        public double dConversionRate;
        public double dSale;
        public double dDiscount;
        public double dSurcharge;
        public double dFastTrack;
        public double dFreight;
        public string sCity;
    };

    protected void Page_Load(object sender, EventArgs e)
    {
        // load all fiscal years
        YearList.Items.Add("2018");
        YearList.Items.Add("2017");
        YearList.Items.Add("2016");
        YearList.Items.Add("2015");
        YearList.Items.Add("2014");
        YearList.Items.Add("2013");

        // force check sapa
        SelectSapa.Checked = true;
    }

    protected void ButtonSave_Click(object sender, EventArgs e)
    {
        if (GridViewSalesReport.Rows.Count > 0)
        {
            // write to excel
            string fileName = @"C:\Trial Balance\" + User.Identity.Name + " Monthly Sales Report for " + sCustomerName + " at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + ".csv";
            if (Directory.Exists(@"C:\Trial Balance\"))
            {
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
            }
            else
            {
                Directory.CreateDirectory(@"C:\Trial Balance\");
            }
            GenerateCSVFile(fileName);
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

    private void GenerateCSVFile(string filePath)
    {
        // create file
        StreamWriter writer = new StreamWriter(filePath);
        // insert title
        writer.WriteLine("Sapa Sales Report for " + sCustomerName + " at " + DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss"));
        writer.WriteLine();
        foreach (TableCell cell in GridViewSalesReport.HeaderRow.Cells)
        {
            writer.Write(cell.Text + ",");
        }
        writer.WriteLine();
        // write content
        foreach (GridViewRow row in GridViewSalesReport.Rows)
        {
            string content = string.Empty;
            foreach (TableCell cell in row.Cells)
            {
                content += "\"" + cell.Text + "\",";
            }
            content = content.Replace("&#39;", "'");
            content = content.Replace("&amp;", "&");
            content = content.Replace("&nbsp;", " ");
            writer.WriteLine(content);
        }
        // save file
        writer.Flush();
        writer.Close();
    }

    protected void ButtonGenerate_Click(object sender, EventArgs e)
    {
        GridViewSalesReport.DataSource = null;
        GridViewSalesReport.Caption = string.Empty;
        GridViewSalesReport.DataBind();
        // get customer name
        if (SelectHydro.Checked)
        {
            sCustomerName = "Hydro";
        }
        else if (SelectSapa.Checked)
        {
            sCustomerName = "Sapa";
        }
        // get data
        List<Invoice> invoiceList = new List<Invoice>();
        ExcoODBC database = ExcoODBC.Instance;
        database.Open(Database.CMSDAT);
        string query = "";
        if (MonthList.SelectedValue == "12")
        {
            query = "select distinct dhinv#, dhexrt, bvcity from cmsdat.oih, cmsdat.cust where (dhbnam like '%" + sCustomerName + "%' or dhbnam like '%" + sCustomerName.ToUpper() + "%') and (dhidat<='" + (Convert.ToInt32(YearList.SelectedValue.ToString()) + 1).ToString() + "-01-01' and dhidat>='" + YearList.SelectedValue + "-" + MonthList.SelectedValue + "-01') and bvcust=dhscs# order by bvcity";
        }
        else
        {
            query = "select distinct dhinv#, dhexrt, bvcity from cmsdat.oih, cmsdat.cust where (dhbnam like '%" + sCustomerName + "%' or dhbnam like '%" + sCustomerName.ToUpper() + "%') and (dhidat<'" + YearList.SelectedValue + "-" + (MonthList.SelectedIndex + 2).ToString("D2") + "-01' and dhidat>='" + YearList.SelectedValue + "-" + MonthList.SelectedValue + "-01') and bvcust=dhscs# order by bvcity";
        }
        OdbcDataReader reader = database.RunQuery(query);
        while (reader.Read())
        {
            Invoice invoice = new Invoice();
            invoice.sInvNum = reader["dhinv#"].ToString().Trim();
            invoice.dConversionRate = 1;//Convert.ToDouble(reader["dhexrt"]);
            invoice.sCity = reader["bvcity"].ToString().Trim();
            invoiceList.Add(invoice);
        }
        reader.Close();
        // sale
        for (int i = 0; i < invoiceList.Count; i++)
        {
            query = "select coalesce(sum(dipric*diqtsp),0.0) from cmsdat.oid where diglcd='SAL' and diinv#=" + invoiceList[i].sInvNum;
            reader = database.RunQuery(query);
            if (reader.Read())
            {
                Invoice invoice = invoiceList[i];
                invoice.dSale = Convert.ToDouble(reader[0]);
                invoiceList[i] = invoice;
            }
            reader.Close();
        }
        // discount
        for (int i = 0; i < invoiceList.Count; i++)
        {
            query = "select coalesce(sum(fldext),0.0) as value from cmsdat.ois where (fldisc like 'D%' or fldisc like 'M%') and flinv#=" + invoiceList[i].sInvNum;
            reader = database.RunQuery(query);
            if (reader.Read())
            {
                Invoice invoice = invoiceList[i];
                invoice.dDiscount = Convert.ToDouble(reader[0]);
                invoiceList[i] = invoice;
            }
            reader.Close();
        }
        // fast track
        for (int i = 0; i < invoiceList.Count; i++)
        {
            query = "select coalesce(sum(fldext),0.0) as value from cmsdat.ois where fldisc like 'F%' and flinv#=" + invoiceList[i].sInvNum;
            reader = database.RunQuery(query);
            if (reader.Read())
            {
                Invoice invoice = invoiceList[i];
                invoice.dFastTrack = Convert.ToDouble(reader[0]);
                invoiceList[i] = invoice;
            }
            reader.Close();
        }
        // surcharge
        for (int i = 0; i < invoiceList.Count; i++)
        {
            query = "select coalesce(sum(fldext),0.0) as value from cmsdat.ois where (fldisc like 'S%' or fldisc like 'P%') and flinv#=" + invoiceList[i].sInvNum;
            reader = database.RunQuery(query);
            if (reader.Read())
            {
                Invoice invoice = invoiceList[i];
                invoice.dSurcharge = Convert.ToDouble(reader[0]);
                invoiceList[i] = invoice;
            }
            reader.Close();
        }
        // freight
        for (int i = 0; i < invoiceList.Count; i++)
        {
            query = "select DIEXT from cmsdat.oid where dipart like 'FREIGHT%' AND DIINV# = " + invoiceList[i].sInvNum;
            reader = database.RunQuery(query);
            if (reader.Read())
            {
                Invoice invoice = invoiceList[i];
                invoice.dFreight = Convert.ToDouble(reader[0]);
                invoiceList[i] = invoice;
            }
            reader.Close();
        }
        // write to data table
        DataTable table = new DataTable();
        double dSummaryByCity = 0.0;
        string sCity = string.Empty;
        List<double> summaryList = new List<double>();
        foreach (Invoice invoice in invoiceList)
        {
            if (sCity != invoice.sCity)
            {
                if (sCity != string.Empty)
                {
                    // write to grid view
                    table.Columns.Add(sCity);
                    summaryList.Add(dSummaryByCity);
                }
                // new city
                dSummaryByCity = (invoice.dSale + invoice.dFastTrack + invoice.dSurcharge + invoice.dFreight + invoice.dDiscount) * invoice.dConversionRate;
                sCity = invoice.sCity;
            }
            else
            {
                dSummaryByCity += (invoice.dSale + invoice.dFastTrack + invoice.dSurcharge + invoice.dFreight + invoice.dDiscount) * invoice.dConversionRate;
            }
        }
        // last city
        table.Columns.Add(sCity);
        summaryList.Add(dSummaryByCity);
        // build row
        DataRow row = table.NewRow();
        for (int i = 0; i < summaryList.Count; i++)
        {
            row[i] = summaryList[i].ToString("C2");
        }
        table.Rows.Add(row);
        // write to grid view
        DateTime datename = new DateTime(Convert.ToInt32(YearList.SelectedValue), Convert.ToInt32(MonthList.SelectedValue), 1);
        string monthname = datename.ToString("MMMM");
        GridViewSalesReport.Caption = "<font size=\"9\" color=\"red\">Monthly Sales Report for " + sCustomerName + " at " + monthname + "/" + YearList.SelectedValue;
        GridViewSalesReport.DataSource = table;
        GridViewSalesReport.DataBind();
    }
}