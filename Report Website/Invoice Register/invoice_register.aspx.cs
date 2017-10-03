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
    protected void Page_Load(object sender, EventArgs e)
    {
        //if (true == (bool)Session["_hasInvoiceRegisterTexasReport"])
        //{
            TextBoxSelectPlant.Visible = true;
            ButtonTexas.Visible = true;
        //}
        //if (true == (bool)Session["_hasInvoiceRegisterColombiaReport"])
        //{
            TextBoxSelectPlant.Visible = true;
            ButtonColombia.Visible = true;
        //}
        //if (true == (bool)Session["_hasInvoiceRegisterMichiganReport"])
        //{
            TextBoxSelectPlant.Visible = true;
            ButtonMichigan.Visible = true;
        //}
        //if (true == (bool)Session["_hasInvoiceRegisterMarkhamReport"])
        //{
            TextBoxSelectPlant.Visible = true;
            ButtonMarkham.Visible = true;
        //}
    }

    protected void ButtonSave_Click(object sender, EventArgs e)
    {
        // write to excel
        string fileName = @"C:\Invoice Register\" + User.Identity.Name + " Invoice Register for " + Session["_plantName"].ToString() + " at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + ".csv";
        if (File.Exists(fileName))
        {
            File.Delete(fileName);
        }
        StreamWriter writer = new StreamWriter(fileName);
        if (GridViewInvoiceRegisterCAD.Visible)
        {
            GenerateCSVFile(ref writer, GridViewInvoiceRegisterCAD);
        }
        if (GridViewInvoiceRegisterPESO.Visible)
        {
            GenerateCSVFile(ref writer, GridViewInvoiceRegisterPESO);
        }
        if (GridViewInvoiceRegisterUSD.Visible)
        {
            GenerateCSVFile(ref writer, GridViewInvoiceRegisterUSD);
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

    protected void ButtonMichigan_Click(object sender, EventArgs e)
    {
        Session["_plantID"] = "003";
        Session["_plantName"] = "Michigan";
        GridView gridView = GridViewInvoiceRegisterUSD;
        gridView.Caption = "<font size=\"9\" color=\"red\">Month To Date Michigan Invoice Register Report (USD): " + DateTime.Now.ToString("MM/dd/yyyy H:mm") + "</font>";
        PanelInvoiceRegisterUSD.Visible = true;
        PanelInvoiceRegisterCAD.Visible = false;
        PanelInvoiceRegisterPESO.Visible = false;
        ExcoCalendar calendar = new ExcoCalendar(DateTime.Now.Year - 2000, DateTime.Now.Month, false, 3);
        // create grid view table
        DataTable report = new DataTable();
        report.Columns.Add("Credit Note");
        report.Columns.Add("Inv#");
        report.Columns.Add("Inv Date");
        report.Columns.Add("SO#");
        report.Columns.Add("Cust Name");
        report.Columns.Add("Cust ID");
        report.Columns.Add("Currency");
        report.Columns.Add("Die Sale");
        report.Columns.Add("Discount");
        report.Columns.Add("Fast Track");
        report.Columns.Add("Surcharge");
        report.Columns.Add("Freight");
        report.Columns.Add("Total");
        report.Columns.Add("Posted");
        // build connection
        OdbcConnection connection = new OdbcConnection(ConfigurationManager.ConnectionStrings["cms1"].ConnectionString);
        connection.Open();
        OdbcCommand command = new OdbcCommand();
        command.Connection = connection;
        string query = "select dhincr, dhinv#, dhidat, dhord#, dhbnam, dhbcs#, dhterr, dhcurr, dhtoti, dhpost from cmsdat.oih where dhplnt='003' and (dharyr=" + calendar.GetFiscalYear().ToString("D2") + " and dharpr=" + calendar.GetFiscalMonth().ToString() + " or (dhidat>='" + (calendar.GetCalendarYear() + 2000).ToString() + "-" + calendar.GetCalendarMonth().ToString("D2") + "-01' and dhidat<=date(current timestamp) and dharyr=0 and dharpr=0)) order by dhinv#";
        command.CommandText = query;
        OdbcDataReader reader = command.ExecuteReader();
        List<DataRow> rowList = new List<DataRow>();
        double totalDieSale = 0.0;
        double totalDiscount = 0.0;
        double totalFastTrack = 0.0;
        double totalSurcharge = 0.0;
        double totalFreight = 0.0;
        double totalTotal = 0.0;
        int invoiceNum = 0;
        // get main data
        while (reader.Read())
        {
            DataRow row = report.NewRow();
            row["Inv#"] = reader["dhinv#"].ToString().Trim();
            invoiceNum = Convert.ToInt32(reader["dhinv#"]);
            row["Credit Note"] = reader["dhincr"].ToString().Trim();
            row["Inv Date"] = Convert.ToDateTime(reader["dhidat"]).ToString("MM/dd/yyyy").Trim();
            row["SO#"] = reader["dhord#"].ToString().Trim();
            row["Cust Name"] = reader["dhbnam"].ToString().Trim();
            row["Cust ID"] = reader["dhbcs#"].ToString().Trim();
            row["Currency"] = reader["dhcurr"].ToString().Trim();
            if (string.Empty == reader["dhtoti"].ToString())
            {
                row["Total"] = "$0.00";
            }
            else
            {
                row["Total"] = Convert.ToDouble(reader["dhtoti"]).ToString("C2");
                totalTotal += Convert.ToDouble(reader["dhtoti"]);
            }
            row["Posted"] = reader["dhpost"].ToString().Trim();
            rowList.Add(row);
        }
        reader.Close();
        // sale
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(dipric*diqtsp) as sale from cmsdat.oid where diglcd='SAL' and diinv#=" + row["Inv#"].ToString();
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    if (string.Empty == reader["sale"].ToString())
                    {
                        row["Die Sale"] = "$0.00";
                    }
                    else
                    {
                        row["Die Sale"] = Convert.ToDouble(reader["sale"]).ToString("C2");
                        totalDieSale += Convert.ToDouble(reader["sale"]);
                    }
                }
                reader.Close();
            }
        }
        // discount
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(fldext) as value from cmsdat.ois where flinv#=" + row["Inv#"].ToString() + " and (fldisc like 'D%' or fldisc like 'M%')";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    if (string.Empty == reader["value"].ToString())
                    {
                        row["Discount"] = "$0.00";
                    }
                    else
                    {
                        row["Discount"] = Convert.ToDouble(reader["value"]).ToString("C2");
                        totalDiscount += Convert.ToDouble(reader["value"]);
                    }
                }
                reader.Close();
            }
        }
        // fast track
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(fldext) as value from cmsdat.ois where flinv#=" + row["Inv#"].ToString() + " and fldisc like 'F%'";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    if (string.Empty == reader["value"].ToString())
                    {
                        row["Fast Track"] = "$0.00";
                    }
                    else
                    {
                        row["Fast Track"] = Convert.ToDouble(reader["value"]).ToString("C2");
                        totalFastTrack += Convert.ToDouble(reader["value"]);
                    }
                }
                reader.Close();
            }
        }
        // surcharge
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(fldext) as value from cmsdat.ois where flinv#=" + row["Inv#"].ToString() + " and (fldisc like 'S%' or fldisc like 'P%')";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    if (string.Empty == reader["value"].ToString())
                    {
                        row["Surcharge"] = "$0.00";
                    }
                    else
                    {
                        row["Surcharge"] = Convert.ToDouble(reader["value"]).ToString("C2");
                        totalSurcharge += Convert.ToDouble(reader["value"]);
                    }
                }
                reader.Close();
            }
        }
        // frieght
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(dipric*diqtsp) as frt from cmsdat.oid where diglcd='FRT' and diinv#=" + row["Inv#"].ToString();
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    if (string.Empty == reader["frt"].ToString())
                    {
                        row["Freight"] = "$0.00";
                    }
                    else
                    {
                        row["Freight"] = Convert.ToDouble(reader["frt"]).ToString("C2");
                        totalFreight += Convert.ToDouble(reader["frt"]);
                    }
                }
                reader.Close();
            }
        }
        // summary row
        DataRow totalRow = report.NewRow();
        totalRow["Currency"] = "Consolidate:";
        totalRow["Die Sale"] = totalDieSale.ToString("C2");
        totalRow["Discount"] = totalDiscount.ToString("C2");
        totalRow["Fast Track"] = totalFastTrack.ToString("C2");
        totalRow["Surcharge"] = totalSurcharge.ToString("C2");
        totalRow["Freight"] = totalFreight.ToString("C2");
        totalRow["Total"] = totalTotal.ToString("C2");
        rowList.Add(totalRow);
        // bind data
        foreach (DataRow row in rowList)
        {
            report.Rows.Add(row);
        }
        report.AcceptChanges();
        gridView.DataSource = report;
        gridView.DataBind();
        // adjust style
        foreach (GridViewRow row in gridView.Rows)
        {
            // not posted
            if ("N" == row.Cells[row.Cells.Count - 1].Text)
            {
                Style style = new Style();
                style.BackColor = Color.Salmon;
                row.ApplyStyle(style);
            }
            // summary row
            else if ("&nbsp;" == row.Cells[1].Text)
            {
                Style style = new Style();
                style.Font.Bold = true;
                row.ApplyStyle(style);
            }
        }
    }

    protected void ButtonMarkham_Click(object sender, EventArgs e)
    {
        Session["_plantID"] = "001";
        Session["_plantName"] = "Markham";
        PanelInvoiceRegisterPESO.Visible = false;
        ExcoCalendar calendar = new ExcoCalendar(DateTime.Now.Year - 2000, DateTime.Now.Month, false, 1);
        // create grid view table
        DataTable report = new DataTable();
        report.Columns.Add("Credit Note");
        report.Columns.Add("Inv#");
        report.Columns.Add("Inv Date");
        report.Columns.Add("SO#");
        report.Columns.Add("Cust Name");
        report.Columns.Add("Cust ID");
        report.Columns.Add("Currency");
        report.Columns.Add("Die Sale");
        report.Columns.Add("Surcharge");
        report.Columns.Add("Fast Track");
        report.Columns.Add("Freight");
        report.Columns.Add("Discount");
        report.Columns.Add("Total Sale");
        report.Columns.Add("Scrap Sale");
        report.Columns.Add("Tax");
        report.Columns.Add("Inv Total");
        report.Columns.Add("Posted");
        // build connection
        OdbcConnection connection = new OdbcConnection(ConfigurationManager.ConnectionStrings["cms1"].ConnectionString);
        connection.Open();
        OdbcCommand command = new OdbcCommand();
        command.Connection = connection;
        string query = "select dhincr, dhinv#, dhidat, dhbnam, dhbcs#, dhord#, dhterr, dhcurr, dhpost, dhtoti from cmsdat.oih where dhplnt='001' and (dharyr=" + calendar.GetFiscalYear().ToString("D2") + " and dharpr=" + calendar.GetFiscalMonth().ToString() + " or (dhidat>='" + (calendar.GetCalendarYear() + 2000).ToString() + "-" + calendar.GetCalendarMonth().ToString("D2") + "-01' and dhidat<=date(current timestamp) and dharyr=0 and dharpr=0)) order by dhinv#";
        command.CommandText = query;
        OdbcDataReader reader = command.ExecuteReader();
        List<DataRow> rowList = new List<DataRow>();
        int invoiceNum = 0;
        // get main data
        while (reader.Read())
        {
            DataRow row = report.NewRow();
            row["Inv#"] = reader["dhinv#"].ToString().Trim();
            invoiceNum = Convert.ToInt32(reader["dhinv#"]);
            row["Credit Note"] = reader["dhincr"].ToString().Trim();
            row["Inv Date"] = Convert.ToDateTime(reader["dhidat"]).ToString("MM/dd/yyyy").Trim();
            row["SO#"] = reader["dhord#"].ToString().Trim();
            row["Cust Name"] = reader["dhbnam"].ToString().Trim();
            row["Cust ID"] = reader["dhbcs#"].ToString().Trim();
            row["Currency"] = reader["dhcurr"].ToString().Trim();
            if (string.Empty == reader["dhtoti"].ToString())
            {
                row["Total"] = "0.00";
            }
            else
            {
                row["Inv Total"] = reader["dhtoti"];
            }
            row["Posted"] = reader["dhpost"].ToString().Trim();
            rowList.Add(row);
        }
        reader.Close();
        // die sale
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(dipric*diqtsp) as sale from cmsdat.oid where diglcd='SAL' and diinv#=" + row["Inv#"].ToString();
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["sale"].ToString())
                {
                    row["Die Sale"] = reader["sale"];
                }
                else
                {
                    row["Die Sale"] = "0.00";
                }
                reader.Close();
            }
        }
        // discount
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(fldext) as value from cmsdat.ois where flinv#=" + row["Inv#"].ToString() + " and (fldisc like 'D%' or fldisc like 'M%')";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["value"].ToString())
                {
                    row["Discount"] = reader["value"];
                }
                else
                {
                    row["Discount"] = "0.00";
                }
                reader.Close();
            }
        }
        // fast track
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(fldext) as value from cmsdat.ois where flinv#=" + row["Inv#"].ToString() + " and fldisc like 'F%'";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["value"].ToString())
                {
                    row["Fast Track"] = reader["value"];
                }
                else
                {
                    row["Fast Track"] = "0.00";
                }
                reader.Close();
            }
        }
        // frieght
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(dipric*diqtsp) as frt from cmsdat.oid where diglcd='FRT' and diinv#=" + row["Inv#"].ToString();
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["frt"].ToString())
                {
                    row["Freight"] = reader["frt"];
                }
                else
                {
                    row["Freight"] = "0.00";
                }
                reader.Close();
            }
        }
        // surcharge
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(fldext) as value from cmsdat.ois where flinv#=" + row["Inv#"].ToString() + " and (fldisc like 'S%' or fldisc like 'P%')";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["value"].ToString())
                {
                    row["Surcharge"] = reader["value"];
                }
                else
                {
                    row["Surcharge"] = "0.00";
                }
                reader.Close();
            }
        }
        // total sale
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                row["Total Sale"] = Convert.ToDouble(row["Die Sale"]) + Convert.ToDouble(row["Surcharge"]) + Convert.ToDouble(row["Fast Track"]) + Convert.ToDouble(row["Freight"]) + Convert.ToDouble(row["Discount"]);
            }
        }
        // tax
        foreach (DataRow row in rowList)
        {
            query = "select notxam from cmsdat.oix where noinv#=" + row["Inv#"].ToString();
            command.CommandText = query;
            reader = command.ExecuteReader();
            if (reader.Read() && string.Empty != reader["notxam"].ToString().Trim())
            {
                row["Tax"] = reader["notxam"];
            }
            else
            {
                row["Tax"] = "0.00";
            }
            reader.Close();
        }
        // bind data to Peso and Usd report
        report.AcceptChanges();
        DataTable reportCad = new DataTable();
        reportCad.Columns.Add("Credit Note");
        reportCad.Columns.Add("Inv#");
        reportCad.Columns.Add("Inv Date");
        reportCad.Columns.Add("SO#");
        reportCad.Columns.Add("Cust Name");
        reportCad.Columns.Add("Cust ID");
        reportCad.Columns.Add("Currency");
        reportCad.Columns.Add("Die Sale");
        reportCad.Columns.Add("Surcharge");
        reportCad.Columns.Add("Fast Track");
        reportCad.Columns.Add("Freight");
        reportCad.Columns.Add("Discount");
        reportCad.Columns.Add("Total Sale");
        reportCad.Columns.Add("Tax");
        reportCad.Columns.Add("Inv Total");
        reportCad.Columns.Add("Posted");
        DataTable reportUsd = new DataTable();
        reportUsd.Columns.Add("Credit Note");
        reportUsd.Columns.Add("Inv#");
        reportUsd.Columns.Add("Inv Date");
        reportUsd.Columns.Add("SO#");
        reportUsd.Columns.Add("Cust Name");
        reportUsd.Columns.Add("Cust ID");
        reportUsd.Columns.Add("Currency");
        reportUsd.Columns.Add("Die Sale");
        reportUsd.Columns.Add("Surcharge");
        reportUsd.Columns.Add("Fast Track");
        reportUsd.Columns.Add("Freight");
        reportUsd.Columns.Add("Discount");
        reportUsd.Columns.Add("Total Sale");
        reportUsd.Columns.Add("Tax");
        reportUsd.Columns.Add("Inv Total");
        reportUsd.Columns.Add("Posted");

        double totalDieSaleCad = 0.0;
        double totalDiscountCad = 0.0;
        double totalFastTrackCad = 0.0;
        double totalSurchargeCad = 0.0;
        double totalFreightCad = 0.0;
        double totalTotalSaleCad = 0.0;
        double totalTaxCad = 0.0;
        double totalInvTotalCad = 0.0;
        double totalDieSaleUsd = 0.0;
        double totalDiscountUsd = 0.0;
        double totalFastTrackUsd = 0.0;
        double totalSurchargeUsd = 0.0;
        double totalFreightUsd = 0.0;
        double totalTotalSaleUsd = 0.0;
        double totalTaxUsd = 0.0;
        double totalInvTotalUsd = 0.0;
        foreach (DataRow row in rowList)
        {
            if ("CA" == row["Currency"].ToString())
            {
                DataRow cadRow = reportCad.NewRow();
                for (int i = 0; i < 8; i++)
                {
                    cadRow[i] = row[i];
                }
                cadRow["Die Sale"] = Convert.ToDouble(row["Die Sale"]).ToString("C2");
                cadRow["Surcharge"] = Convert.ToDouble(row["Surcharge"]).ToString("C2");
                cadRow["Fast Track"] = Convert.ToDouble(row["Fast Track"]).ToString("C2");
                cadRow["Freight"] = Convert.ToDouble(row["Freight"]).ToString("C2");
                cadRow["Discount"] = Convert.ToDouble(row["Discount"]).ToString("C2");
                cadRow["Total Sale"] = Convert.ToDouble(row["Total Sale"]).ToString("C2");
                cadRow["Tax"] = Convert.ToDouble(row["Tax"]).ToString("C2");
                cadRow["Inv Total"] = Convert.ToDouble(row["Inv Total"]).ToString("C2");
                cadRow["Posted"] = row["Posted"];
                reportCad.Rows.Add(cadRow);
                totalDieSaleCad += Convert.ToDouble(row["Die Sale"]);
                totalDiscountCad += Convert.ToDouble(row["Surcharge"].ToString());
                totalFastTrackCad += Convert.ToDouble(row["Fast Track"].ToString());
                totalSurchargeCad += Convert.ToDouble(row["Freight"].ToString());
                totalFreightCad += Convert.ToDouble(row["Discount"].ToString());
                totalTotalSaleCad += Convert.ToDouble(row["Total Sale"].ToString());
                totalTaxCad += Convert.ToDouble(row["Tax"].ToString());
                totalInvTotalCad += Convert.ToDouble(row["Inv Total"].ToString());
            }
            else if ("US" == row["Currency"].ToString())
            {
                DataRow usdRow = reportUsd.NewRow();
                for (int i = 0; i < 8; i++)
                {
                    usdRow[i] = row[i];
                }
                usdRow["Die Sale"] = Convert.ToDouble(row["Die Sale"]).ToString("C2");
                usdRow["Surcharge"] = Convert.ToDouble(row["Surcharge"]).ToString("C2");
                usdRow["Fast Track"] = Convert.ToDouble(row["Fast Track"]).ToString("C2");
                usdRow["Freight"] = Convert.ToDouble(row["Freight"]).ToString("C2");
                usdRow["Discount"] = Convert.ToDouble(row["Discount"]).ToString("C2");
                usdRow["Total Sale"] = Convert.ToDouble(row["Total Sale"]).ToString("C2");
                usdRow["Tax"] = Convert.ToDouble(row["Tax"]).ToString("C2");
                usdRow["Inv Total"] = Convert.ToDouble(row["Inv Total"]).ToString("C2");
                usdRow["Posted"] = row["Posted"];
                reportUsd.Rows.Add(usdRow);
                totalDieSaleUsd += Convert.ToDouble(row["Die Sale"].ToString());
                totalDiscountUsd += Convert.ToDouble(row["Surcharge"].ToString());
                totalFastTrackUsd += Convert.ToDouble(row["Fast Track"].ToString());
                totalSurchargeUsd += Convert.ToDouble(row["Freight"].ToString());
                totalFreightUsd += Convert.ToDouble(row["Discount"].ToString());
                totalTotalSaleUsd += Convert.ToDouble(row["Total Sale"].ToString());
                totalTaxUsd += Convert.ToDouble(row["Tax"].ToString());
                totalInvTotalUsd += Convert.ToDouble(row["Inv Total"].ToString());
            }
            else
            {
                query = "select dhcurr from cmsdat.oih where dhinv#=" + row["Inv#"].ToString();
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && "US" == reader[0].ToString())
                {
                    DataRow usdRow = reportUsd.NewRow();
                    for (int i = 0; i < reportUsd.Columns.Count; i++)
                    {
                        usdRow[i] = row[i];
                    }
                    reportUsd.Rows.Add(usdRow);
                }
                else
                {
                    DataRow cadRow = reportCad.NewRow();
                    for (int i = 0; i < reportCad.Columns.Count; i++)
                    {
                        cadRow[i] = row[i];
                    }
                    reportCad.Rows.Add(cadRow);
                }
                reader.Close();
            }
        }
        // summary row
        DataRow totalRowCad = reportCad.NewRow();
        totalRowCad["Currency"] = "Consolidate:";
        totalRowCad["Die Sale"] = totalDieSaleCad.ToString("C2");
        totalRowCad["Discount"] = totalDiscountCad.ToString("C2");
        totalRowCad["Fast Track"] = totalFastTrackCad.ToString("C2");
        totalRowCad["Surcharge"] = totalSurchargeCad.ToString("C2");
        totalRowCad["Freight"] = totalFreightCad.ToString("C2");
        totalRowCad["Total Sale"] = totalTotalSaleCad.ToString("C2");
        totalRowCad["Tax"] = totalTaxCad.ToString("C2");
        totalRowCad["Inv Total"] = totalInvTotalCad.ToString("C2");
        reportCad.Rows.Add(totalRowCad);
        DataRow totalRowUsd = reportUsd.NewRow();
        totalRowUsd["Currency"] = "Consolidate:";
        totalRowUsd["Die Sale"] = totalDieSaleUsd.ToString("C2");
        totalRowUsd["Discount"] = totalDiscountUsd.ToString("C2");
        totalRowUsd["Fast Track"] = totalFastTrackUsd.ToString("C2");
        totalRowUsd["Surcharge"] = totalSurchargeUsd.ToString("C2");
        totalRowUsd["Freight"] = totalFreightUsd.ToString("C2");
        totalRowUsd["Total Sale"] = totalTotalSaleUsd.ToString("C2");
        totalRowUsd["Tax"] = totalTaxUsd.ToString("C2");
        totalRowUsd["Inv Total"] = totalInvTotalUsd.ToString("C2");
        reportUsd.Rows.Add(totalRowUsd);
        // bind data
        reportCad.AcceptChanges();
        if (reportCad.Rows.Count > 0)
        {
            PanelInvoiceRegisterCAD.Visible = true;
            GridViewInvoiceRegisterCAD.DataSource = reportCad;
            GridViewInvoiceRegisterCAD.DataBind();
            GridViewInvoiceRegisterCAD.Caption = "<font size=\"9\" color=\"red\">Month To Date Markham Invoice Register Report (CAD): " + DateTime.Now.ToString("MM/dd/yyyy H:mm") + "</font>";
        }
        reportUsd.AcceptChanges();
        if (reportUsd.Rows.Count > 0)
        {
            PanelInvoiceRegisterUSD.Visible = true;
            GridViewInvoiceRegisterUSD.DataSource = reportUsd;
            GridViewInvoiceRegisterUSD.DataBind();
            GridViewInvoiceRegisterUSD.Caption = "<font size=\"9\" color=\"red\">Month To Date Markham Invoice Register Report (USD): " + DateTime.Now.ToString("MM/dd/yyyy H:mm") + "</font>";
        }
        // adjust style
        foreach (GridViewRow row in GridViewInvoiceRegisterCAD.Rows)
        {
            // not posted
            if ("N" == row.Cells[row.Cells.Count - 1].Text)
            {
                Style style = new Style();
                style.BackColor = Color.Salmon;
                row.ApplyStyle(style);
            }
            // summary row
            else if ("&nbsp;" == row.Cells[1].Text)
            {
                Style style = new Style();
                style.Font.Bold = true;
                row.ApplyStyle(style);
            }
        }
        foreach (GridViewRow row in GridViewInvoiceRegisterUSD.Rows)
        {
            // not posted
            if ("N" == row.Cells[row.Cells.Count - 1].Text)
            {
                Style style = new Style();
                style.BackColor = Color.Salmon;
                row.ApplyStyle(style);
            }
            // summary row
            else if ("&nbsp;" == row.Cells[1].Text)
            {
                Style style = new Style();
                style.Font.Bold = true;
                row.ApplyStyle(style);
            }
        }
    }

    protected void ButtonColombia_Click(object sender, EventArgs e)
    {
        Session["_plantID"] = "004";
        Session["_plantName"] = "Colombia";
        PanelInvoiceRegisterCAD.Visible = false;
        ExcoCalendar calendar = new ExcoCalendar(DateTime.Now.Year - 2000, DateTime.Now.Month, false, 4);
        // create grid view table
        DataTable report = new DataTable();
        report.Columns.Add("Credit Note");
        report.Columns.Add("Inv#");
        report.Columns.Add("Inv Date");
        report.Columns.Add("SO#");
        report.Columns.Add("Cust Name");
        report.Columns.Add("Cust ID");
        report.Columns.Add("Currency");
        report.Columns.Add("Die Sale");
        report.Columns.Add("Surcharge");
        report.Columns.Add("Fast Track");
        report.Columns.Add("Freight");
        report.Columns.Add("Discount");
        report.Columns.Add("Total Sale");
        report.Columns.Add("Scrap Sale");
        report.Columns.Add("RET 3.5%");
        report.Columns.Add("IVA RET 8%");
        report.Columns.Add("IVA 16%");
        report.Columns.Add("Inv Total");
        report.Columns.Add("Posted");



        /*
        var connection = new OdbcConnection(ConfigurationManager.ConnectionStrings["cms1"].ConnectionString);
        connection.Open();
        var command = new OdbcCommand();
        command.Connection = connection;
        var query = "select dhincr, dhinv#, dhidat, dhord#, dhbnam, dhbcs#, dhterr, dhcurr, dhtoti, dhpost from cmsdat.oih where dhplnt='004' and (dharyr=" + calendar.GetFiscalYear().ToString("D2") + " and dharpr=" + calendar.GetFiscalMonth().ToString() + " or (dhidat>='" + (calendar.GetCalendarYear() + 2000).ToString() + "-" + calendar.GetCalendarMonth().ToString("D2") + "-01' and dhidat<=date(current timestamp) and dharyr=0 and dharpr=0)) order by dhinv#";
        command.CommandText = query;
        var reader = command.ExecuteReader();
         * */
        // build connection
        OdbcConnection connection = new OdbcConnection(ConfigurationManager.ConnectionStrings["cms1"].ConnectionString);
        connection.Open();
        OdbcCommand command = new OdbcCommand();
        command.Connection = connection;
        string query = "select dhincr, dhinv#, dhidat, dhord#, dhbnam, dhbcs#, dhterr, dhcurr, dhtoti, dhpost from cmsdat.oih where dhplnt='004' and (dharyr=" + calendar.GetFiscalYear().ToString("D2") + " and dharpr=" + calendar.GetFiscalMonth().ToString() + " or (dhidat>='" + (calendar.GetCalendarYear() + 2000).ToString() + "-" + calendar.GetCalendarMonth().ToString("D2") + "-01' and dhidat<=date(current timestamp) and dharyr=0 and dharpr=0)) order by dhinv#";
        command.CommandText = query;
        OdbcDataReader reader = command.ExecuteReader();
        List<DataRow> rowList = new List<DataRow>();
        int invoiceNum = 0;
        // get main data
        while (reader.Read())
        {
            DataRow row = report.NewRow();
            row["Inv#"] = reader["dhinv#"].ToString().Trim();
            invoiceNum = Convert.ToInt32(reader["dhinv#"]);
            row["Credit Note"] = reader["dhincr"].ToString().Trim();
            row["Inv Date"] = Convert.ToDateTime(reader["dhidat"]).ToString("MM/dd/yyyy").Trim();
            row["SO#"] = reader["dhord#"].ToString().Trim();
            row["Cust Name"] = reader["dhbnam"].ToString().Trim();
            row["Cust ID"] = reader["dhbcs#"].ToString().Trim();
            row["Currency"] = reader["dhcurr"].ToString().Trim();
            if (string.Empty == reader["dhtoti"].ToString())
            {
                row["Total"] = "0.00";
            }
            else
            {
                row["Inv Total"] = reader["dhtoti"];
            }
            row["Posted"] = reader["dhpost"].ToString().Trim();
            rowList.Add(row);
        }
        reader.Close();
        // die sale
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(dipric*diqtsp) as sale from cmsdat.oid where diglcd='SAL' and diinv#=" + row["Inv#"].ToString();
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["sale"].ToString())
                {
                    row["Die Sale"] = reader["sale"];
                }
                else
                {
                    row["Die Sale"] = "0.00";
                }
                reader.Close();
            }
        }
        // discount
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(fldext) as value from cmsdat.ois where flinv#=" + row["Inv#"].ToString() + " and (fldisc like 'D%' or fldisc like 'M%')";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["value"].ToString())
                {
                    row["Discount"] = reader["value"];
                }
                else
                {
                    row["Discount"] = "0.00";
                }
                reader.Close();
            }
        }
        // fast track
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(fldext) as value from cmsdat.ois where flinv#=" + row["Inv#"].ToString() + " and fldisc like 'F%'";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["value"].ToString())
                {
                    row["Fast Track"] = reader["value"];
                }
                else
                {
                    row["Fast Track"] = "0.00";
                }
                reader.Close();
            }
        }
        // frieght
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(dipric*diqtsp) as frt from cmsdat.oid where diglcd='FRT' and diinv#=" + row["Inv#"].ToString();
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["frt"].ToString())
                {
                    row["Freight"] = reader["frt"];
                }
                else
                {
                    row["Freight"] = "0.00";
                }
                reader.Close();
            }
        }
        // surcharge
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(fldext) as value from cmsdat.ois where flinv#=" + row["Inv#"].ToString() + " and (fldisc like 'S%' or fldisc like 'P%')";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["value"].ToString())
                {
                    row["Surcharge"] = reader["value"];
                }
                else
                {
                    row["Surcharge"] = "0.00";
                }
                reader.Close();
            }
        }
        // total sale
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                row["Total Sale"] = Convert.ToDouble(row["Die Sale"]) + Convert.ToDouble(row["Surcharge"]) + Convert.ToDouble(row["Fast Track"]) + Convert.ToDouble(row["Freight"]) + Convert.ToDouble(row["Discount"]);
            }
        }
        // scarp sale
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(dipric*diqtsp) as scr from cmsdat.oid where diglcd='SCR' and diinv#=" + row["Inv#"].ToString();
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["scr"].ToString())
                {
                    row["Scrap Sale"] = reader["scr"];
                }
                else
                {
                    row["Scrap Sale"] = "0.00";
                }
                reader.Close();
            }
        }

        // tax
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString().Trim() && "CP" == row["Currency"].ToString())
            {
                // RET 3.5%
                query = "select notxam from cmsdat.oix where noinv#=" + row["Inv#"].ToString() + " and notxtp=103";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["notxam"].ToString().Trim())
                {
                    row["RET 3.5%"] = reader["notxam"];
                }
                else
                {
                    row["RET 3.5%"] = "0.00";
                }
                reader.Close();
                // IVA RET 8%
                query = "select notxam from cmsdat.oix where noinv#=" + row["Inv#"].ToString() + " and notxtp=102";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["notxam"].ToString())
                {
                    row["IVA RET 8%"] = reader["notxam"];
                }
                else
                {
                    row["IVA RET 8%"] = "0.00";
                }
                reader.Close();
                // IVA 16%
                query = "select notxam from cmsdat.oix where noinv#=" + row["Inv#"].ToString() + " and notxtp=101";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && string.Empty != reader["notxam"].ToString())
                {
                    row["IVA 16%"] = reader["notxam"];
                }
                else
                {
                    row["IVA 16%"] = "0.00";
                }
                reader.Close();
            }
        }
        // bind data to Peso and Usd report
        report.AcceptChanges();
        DataTable reportPeso = new DataTable();
        reportPeso.Columns.Add("Credit Note");
        reportPeso.Columns.Add("Inv#");
        reportPeso.Columns.Add("Inv Date");
        reportPeso.Columns.Add("SO#");
        reportPeso.Columns.Add("Cust Name");
        reportPeso.Columns.Add("Cust ID");
        reportPeso.Columns.Add("Currency");
        reportPeso.Columns.Add("Die Sale");
        reportPeso.Columns.Add("Surcharge");
        reportPeso.Columns.Add("Fast Track");
        reportPeso.Columns.Add("Freight");
        reportPeso.Columns.Add("Discount");
        reportPeso.Columns.Add("Total Sale");
        reportPeso.Columns.Add("Scrap Sale");
        reportPeso.Columns.Add("RET 3.5%");
        reportPeso.Columns.Add("IVA RET 8%");
        reportPeso.Columns.Add("IVA 16%");
        reportPeso.Columns.Add("Inv Total");
        reportPeso.Columns.Add("Posted");
        DataTable reportUsd = new DataTable();
        reportUsd.Columns.Add("Credit Note");
        reportUsd.Columns.Add("Inv#");
        reportUsd.Columns.Add("Inv Date");
        reportUsd.Columns.Add("SO#");
        reportUsd.Columns.Add("Cust Name");
        reportUsd.Columns.Add("Cust ID");
        reportUsd.Columns.Add("Currency");
        reportUsd.Columns.Add("Die Sale");
        reportUsd.Columns.Add("Surcharge");
        reportUsd.Columns.Add("Fast Track");
        reportUsd.Columns.Add("Freight");
        reportUsd.Columns.Add("Discount");
        reportUsd.Columns.Add("Total Sale");
        reportUsd.Columns.Add("Scrap Sale");
        reportUsd.Columns.Add("Inv Total");
        reportUsd.Columns.Add("Posted");

        double totalDieSalePeso = 0.0;
        double totalDiscountPeso = 0.0;
        double totalFastTrackPeso = 0.0;
        double totalSurchargePeso = 0.0;
        double totalFreightPeso = 0.0;
        double totalTotalSalePeso = 0.0;
        double totalScarpSalePeso = 0.0;
        double totalRetPeso = 0.0;
        double totalIvaRetPeso = 0.0;
        double totalIvaPeso = 0.0;
        double totalInvTotalPeso = 0.0;
        double totalDieSaleUsd = 0.0;
        double totalDiscountUsd = 0.0;
        double totalFastTrackUsd = 0.0;
        double totalSurchargeUsd = 0.0;
        double totalFreightUsd = 0.0;
        double totalTotalSaleUsd = 0.0;
        double totalScarpSaleUsd = 0.0;
        double totalInvTotalUsd = 0.0;
        foreach (DataRow row in rowList)
        {
            if ("CP" == row["Currency"].ToString())
            {
                DataRow pesoRow = reportPeso.NewRow();
                for (int i = 0; i < 8; i++)
                {
                    pesoRow[i] = row[i];
                }
                pesoRow["Die Sale"] = Convert.ToDouble(row["Die Sale"]).ToString("C2");
                pesoRow["Surcharge"] = Convert.ToDouble(row["Surcharge"]).ToString("C2");
                pesoRow["Fast Track"] = Convert.ToDouble(row["Fast Track"]).ToString("C2");
                pesoRow["Freight"] = Convert.ToDouble(row["Freight"]).ToString("C2");
                pesoRow["Discount"] = Convert.ToDouble(row["Discount"]).ToString("C2");
                pesoRow["Total Sale"] = Convert.ToDouble(row["Total Sale"]).ToString("C2");
                pesoRow["Scrap Sale"] = Convert.ToDouble(row["Scrap Sale"]).ToString("C2");
                pesoRow["RET 3.5%"] = Convert.ToDouble(row["RET 3.5%"]).ToString("C2");
                pesoRow["IVA RET 8%"] = Convert.ToDouble(row["IVA RET 8%"]).ToString("C2");
                pesoRow["IVA 16%"] = Convert.ToDouble(row["IVA 16%"]).ToString("C2");
                pesoRow["Inv Total"] = Convert.ToDouble(row["Inv Total"]).ToString("C2");
                pesoRow["Posted"] = row["Posted"];
                reportPeso.Rows.Add(pesoRow);
                totalDieSalePeso += Convert.ToDouble(row["Die Sale"]);
                totalDiscountPeso += Convert.ToDouble(row["Surcharge"].ToString());
                totalFastTrackPeso += Convert.ToDouble(row["Fast Track"].ToString());
                totalSurchargePeso += Convert.ToDouble(row["Freight"].ToString());
                totalFreightPeso += Convert.ToDouble(row["Discount"].ToString());
                totalTotalSalePeso += Convert.ToDouble(row["Total Sale"].ToString());
                totalScarpSalePeso += Convert.ToDouble(row["Scrap Sale"].ToString());
                totalRetPeso += Convert.ToDouble(row["RET 3.5%"].ToString());
                totalIvaRetPeso += Convert.ToDouble(row["IVA RET 8%"].ToString());
                totalIvaPeso += Convert.ToDouble(row["IVA 16%"].ToString());
                totalInvTotalPeso += Convert.ToDouble(row["Inv Total"].ToString());
            }
            else if ("US" == row["Currency"].ToString())
            {
                DataRow usdRow = reportUsd.NewRow();
                for (int i = 0; i < 8; i++)
                {
                    usdRow[i] = row[i];
                }
                usdRow["Die Sale"] = Convert.ToDouble(row["Die Sale"]).ToString("C2");
                usdRow["Surcharge"] = Convert.ToDouble(row["Surcharge"]).ToString("C2");
                usdRow["Fast Track"] = Convert.ToDouble(row["Fast Track"]).ToString("C2");
                usdRow["Freight"] = Convert.ToDouble(row["Freight"]).ToString("C2");
                usdRow["Discount"] = Convert.ToDouble(row["Discount"]).ToString("C2");
                usdRow["Total Sale"] = Convert.ToDouble(row["Total Sale"]).ToString("C2");
                usdRow["Scrap Sale"] = Convert.ToDouble(row["Scrap Sale"]).ToString("C2");
                usdRow["Inv Total"] = Convert.ToDouble(row["Inv Total"]).ToString("C2");
                usdRow["Posted"] = row["Posted"];
                reportUsd.Rows.Add(usdRow);
                totalDieSaleUsd += Convert.ToDouble(row["Die Sale"].ToString());
                totalDiscountUsd += Convert.ToDouble(row["Surcharge"].ToString());
                totalFastTrackUsd += Convert.ToDouble(row["Fast Track"].ToString());
                totalSurchargeUsd += Convert.ToDouble(row["Freight"].ToString());
                totalFreightUsd += Convert.ToDouble(row["Discount"].ToString());
                totalTotalSaleUsd += Convert.ToDouble(row["Total Sale"].ToString());
                totalScarpSaleUsd += Convert.ToDouble(row["Scrap Sale"].ToString());
                totalInvTotalUsd += Convert.ToDouble(row["Inv Total"].ToString());
            }
            else
            {
                query = "select dhcurr from cmsdat.oih where dhinv#=" + row["Inv#"].ToString();
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read() && "US" == reader[0].ToString())
                {
                    DataRow usdRow = reportUsd.NewRow();
                    for (int i = 0; i < reportUsd.Columns.Count; i++)
                    {
                        usdRow[i] = row[i];
                    }
                    reportUsd.Rows.Add(usdRow);
                }
                else
                {
                    DataRow pesoRow = reportPeso.NewRow();
                    for (int i = 0; i < reportPeso.Columns.Count; i++)
                    {
                        pesoRow[i] = row[i];
                    }
                    reportPeso.Rows.Add(pesoRow);
                }
                reader.Close();
            }
        }
        // summary row
        DataRow totalRowPeso = reportPeso.NewRow();
        totalRowPeso["Currency"] = "Consolidate:";
        totalRowPeso["Die Sale"] = totalDieSalePeso.ToString("C2");
        totalRowPeso["Discount"] = totalDiscountPeso.ToString("C2");
        totalRowPeso["Fast Track"] = totalFastTrackPeso.ToString("C2");
        totalRowPeso["Surcharge"] = totalSurchargePeso.ToString("C2");
        totalRowPeso["Freight"] = totalFreightPeso.ToString("C2");
        totalRowPeso["Total Sale"] = totalTotalSalePeso.ToString("C2");
        totalRowPeso["Scrap Sale"] = totalScarpSalePeso.ToString("C2");
        totalRowPeso["RET 3.5%"] = totalRetPeso.ToString("C2");
        totalRowPeso["IVA RET 8%"] = totalIvaRetPeso.ToString("C2");
        totalRowPeso["IVA 16%"] = totalIvaPeso.ToString("C2");
        totalRowPeso["Inv Total"] = totalInvTotalPeso.ToString("C2");
        reportPeso.Rows.Add(totalRowPeso);
        DataRow totalRowUsd = reportUsd.NewRow();
        totalRowUsd["Currency"] = "Consolidate:";
        totalRowUsd["Die Sale"] = totalDieSaleUsd.ToString("C2");
        totalRowUsd["Discount"] = totalDiscountUsd.ToString("C2");
        totalRowUsd["Fast Track"] = totalFastTrackUsd.ToString("C2");
        totalRowUsd["Surcharge"] = totalSurchargeUsd.ToString("C2");
        totalRowUsd["Freight"] = totalFreightUsd.ToString("C2");
        totalRowUsd["Total Sale"] = totalTotalSaleUsd.ToString("C2");
        totalRowUsd["Scrap Sale"] = totalScarpSaleUsd.ToString("C2");
        totalRowUsd["Inv Total"] = totalInvTotalUsd.ToString("C2");
        reportUsd.Rows.Add(totalRowUsd);
        // bind data
        reportPeso.AcceptChanges();
        if (reportPeso.Rows.Count > 0)
        {
            PanelInvoiceRegisterPESO.Visible = true;
            GridViewInvoiceRegisterPESO.DataSource = reportPeso;
            GridViewInvoiceRegisterPESO.DataBind();
            GridViewInvoiceRegisterPESO.Caption = "<font size=\"9\" color=\"red\">Month To Date Colombia Invoice Register Report (PESO): " + DateTime.Now.ToString("MM/dd/yyyy H:mm") + "</font>";
        }
        reportUsd.AcceptChanges();
        if (reportUsd.Rows.Count > 0)
        {
            PanelInvoiceRegisterUSD.Visible = true;
            GridViewInvoiceRegisterUSD.DataSource = reportUsd;
            GridViewInvoiceRegisterUSD.DataBind();
            GridViewInvoiceRegisterUSD.Caption = "<font size=\"9\" color=\"red\">Month To Date Colombia Invoice Register Report (USD): " + DateTime.Now.ToString("MM/dd/yyyy H:mm") + "</font>";
        }
        // adjust style
        foreach (GridViewRow row in GridViewInvoiceRegisterPESO.Rows)
        {
            // not posted
            if ("N" == row.Cells[row.Cells.Count - 1].Text)
            {
                Style style = new Style();
                style.BackColor = Color.Salmon;
                row.ApplyStyle(style);
            }
            // summary row
            else if ("&nbsp;" == row.Cells[1].Text)
            {
                Style style = new Style();
                style.Font.Bold = true;
                row.ApplyStyle(style);
            }
        }
        foreach (GridViewRow row in GridViewInvoiceRegisterUSD.Rows)
        {
            // not posted
            if ("N" == row.Cells[row.Cells.Count - 1].Text)
            {
                Style style = new Style();
                style.BackColor = Color.Salmon;
                row.ApplyStyle(style);
            }
            // summary row
            else if ("&nbsp;" == row.Cells[1].Text)
            {
                Style style = new Style();
                style.Font.Bold = true;
                row.ApplyStyle(style);
            }
        }
    }

    protected void ButtonTexas_Click(object sender, EventArgs e)
    {
        Session["_plantID"] = "005";
        Session["_plantName"] = "Texas";
        GridView gridView = GridViewInvoiceRegisterUSD;
        gridView.Caption = "<font size=\"9\" color=\"red\">Month To Date Texas Invoice Register Report (USD): " + DateTime.Now.ToString("MM/dd/yyyy H:mm") + "</font>";
        PanelInvoiceRegisterUSD.Visible = true;
        PanelInvoiceRegisterCAD.Visible = false;
        PanelInvoiceRegisterPESO.Visible = false;
        ExcoCalendar calendar = new ExcoCalendar(DateTime.Now.Year - 2000, DateTime.Now.Month, false, 5);
        // create grid view table
        DataTable report = new DataTable();
        report.Columns.Add("Credit Note");
        report.Columns.Add("Inv#");
        report.Columns.Add("Inv Date");
        report.Columns.Add("SO#");
        report.Columns.Add("Cust Name");
        report.Columns.Add("Cust ID");
        report.Columns.Add("Currency");
        report.Columns.Add("Die Sale");
        report.Columns.Add("Discount");
        report.Columns.Add("Fast Track");
        report.Columns.Add("Surcharge");
        report.Columns.Add("Freight");
        report.Columns.Add("Total");
        report.Columns.Add("Posted");
        // build connection
        OdbcConnection connection = new OdbcConnection(ConfigurationManager.ConnectionStrings["cms1"].ConnectionString); 
        connection.Open();
        OdbcCommand command = new OdbcCommand();
        command.Connection = connection;
        string query = "select dhincr, dhinv#, dhidat, dhord#, dhbnam, dhbcs#, dhterr, dhcurr, dhtoti, dhpost from cmsdat.oih where dhplnt='005' and (dharyr=" + calendar.GetFiscalYear().ToString("D2") + " and dharpr=" + calendar.GetFiscalMonth().ToString() + " or (dhidat>='" + (calendar.GetCalendarYear() + 2000).ToString() + "-" + calendar.GetCalendarMonth().ToString("D2") + "-01' and dhidat<=date(current timestamp) and dharyr=0 and dharpr=0)) order by dhinv#";
        command.CommandText = query;
        OdbcDataReader reader = command.ExecuteReader();
        List<DataRow> rowList = new List<DataRow>();
        double totalDieSale = 0.0;
        double totalDiscount = 0.0;
        double totalFastTrack = 0.0;
        double totalSurcharge = 0.0;
        double totalFreight = 0.0;
        double totalTotal = 0.0;
        int invoiceNum = 0;
        // get main data
        while (reader.Read())
        {
            DataRow row = report.NewRow();
            row["Inv#"] = reader["dhinv#"].ToString().Trim();
            invoiceNum++;
            row["Credit Note"] = reader["dhincr"].ToString().Trim();
            row["Inv Date"] = Convert.ToDateTime(reader["dhidat"]).ToString("MM/dd/yyyy").Trim();
            row["SO#"] = reader["dhord#"].ToString().Trim();
            row["Cust Name"] = reader["dhbnam"].ToString().Trim();
            row["Cust ID"] = reader["dhbcs#"].ToString().Trim();
            row["Currency"] = reader["dhcurr"].ToString().Trim();
            if (string.Empty == reader["dhtoti"].ToString())
            {
                row["Total"] = "$0.00";
            }
            else
            {
                row["Total"] = Convert.ToDouble(reader["dhtoti"]).ToString("C2");
                totalTotal += Convert.ToDouble(reader["dhtoti"]);
            }
            row["Posted"] = reader["dhpost"].ToString().Trim();
            rowList.Add(row);
        }
        reader.Close();
        // sale
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(dipric*diqtsp) as sale from cmsdat.oid where diglcd='SAL' and diinv#=" + row["Inv#"].ToString();
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    if (string.Empty == reader["sale"].ToString())
                    {
                        row["Die Sale"] = "$0.00";
                    }
                    else
                    {
                        row["Die Sale"] = Convert.ToDouble(reader["sale"]).ToString("C2");
                        totalDieSale += Convert.ToDouble(reader["sale"]);
                    }
                }
                reader.Close();
            }
        }
        // discount
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(fldext) as value from cmsdat.ois where flinv#=" + row["Inv#"].ToString() + " and (fldisc like 'D%' or fldisc like 'M%')";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    if (string.Empty == reader["value"].ToString())
                    {
                        row["Discount"] = "$0.00";
                    }
                    else
                    {
                        row["Discount"] = Convert.ToDouble(reader["value"]).ToString("C2");
                        totalDiscount += Convert.ToDouble(reader["value"]);
                    }
                }
                reader.Close();
            }
        }
        // fast track
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(fldext) as value from cmsdat.ois where flinv#=" + row["Inv#"].ToString() + " and fldisc like 'F%'";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    if (string.Empty == reader["value"].ToString())
                    {
                        row["Fast Track"] = "$0.00";
                    }
                    else
                    {
                        row["Fast Track"] = Convert.ToDouble(reader["value"]).ToString("C2");
                        totalFastTrack += Convert.ToDouble(reader["value"]);
                    }
                }
                reader.Close();
            }
        }
        // surcharge
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(fldext) as value from cmsdat.ois where flinv#=" + row["Inv#"].ToString() + " and (fldisc like 'S%' or fldisc like 'P%')";
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    if (string.Empty == reader["value"].ToString())
                    {
                        row["Surcharge"] = "$0.00";
                    }
                    else
                    {
                        row["Surcharge"] = Convert.ToDouble(reader["value"]).ToString("C2");
                        totalSurcharge += Convert.ToDouble(reader["value"]);
                    }
                }
                reader.Close();
            }
        }
        // frieght
        foreach (DataRow row in rowList)
        {
            if (string.Empty != row["Credit Note"].ToString())
            {
                query = "select sum(dipric*diqtsp) as frt from cmsdat.oid where diglcd='FRT' and diinv#=" + row["Inv#"].ToString();
                command.CommandText = query;
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    if (string.Empty == reader["frt"].ToString())
                    {
                        row["Freight"] = "$0.00";
                    }
                    else
                    {
                        row["Freight"] = Convert.ToDouble(reader["frt"]).ToString("C2");
                        totalFreight += Convert.ToDouble(reader["frt"]);
                    }
                }
                reader.Close();
            }
        }
        // summary row
        DataRow totalRow = report.NewRow();
        totalRow["Currency"] = "Consolidate:";
        totalRow["Die Sale"] = totalDieSale.ToString("C2");
        totalRow["Discount"] = totalDiscount.ToString("C2");
        totalRow["Fast Track"] = totalFastTrack.ToString("C2");
        totalRow["Surcharge"] = totalSurcharge.ToString("C2");
        totalRow["Freight"] = totalFreight.ToString("C2");
        totalRow["Total"] = totalTotal.ToString("C2");
        rowList.Add(totalRow);
        // bind data
        foreach (DataRow row in rowList)
        {
            report.Rows.Add(row);
        }
        report.AcceptChanges();
        gridView.DataSource = report;
        gridView.DataBind();
        // adjust style
        foreach (GridViewRow row in gridView.Rows)
        {
            // not posted
            if ("N" == row.Cells[row.Cells.Count - 1].Text)
            {
                Style style = new Style();
                style.BackColor = Color.Salmon;
                row.ApplyStyle(style);
            }
            // summary row
            else if ("&nbsp;" == row.Cells[1].Text)
            {
                Style style = new Style();
                style.Font.Bold = true;
                row.ApplyStyle(style);
            }
        }
    }
}