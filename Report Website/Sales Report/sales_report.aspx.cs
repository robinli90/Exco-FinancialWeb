using System;
using System.IO;
using System.Drawing;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.Odbc;
using System.Configuration;
using ExcoUtility;
using SalesReport;

public partial class sales_report : System.Web.UI.Page
{
    static private int plantID = 0;

    protected void Page_Load(object sender, EventArgs e)
    {
        // process user privilege
        if (true == (bool)Session["_hasSalesReportTexas"])
        {
            ButtonTexas.Visible = true;
        }
        else
        {
            ButtonTexas.Visible = false;
        }
        if (true == (bool)Session["_hasSalesReportColombia"])
        {
            ButtonColombia.Visible = true;
        }
        else
        {
            ButtonColombia.Visible = false;
        }
        if (true == (bool)Session["_hasSalesReportMichigan"])
        {
            ButtonMichigan.Visible = true;
        }
        else
        {
            ButtonMichigan.Visible = false;
        }
        if (true == (bool)Session["_hasSalesReportMarkham"])
        {
            ButtonMarkham.Visible = true;
        }
        else
        {
            ButtonMarkham.Visible = false;
        }

        // load all fiscal years
        var connection = new OdbcConnection(ConfigurationManager.ConnectionStrings["cms1"].ConnectionString);
        connection.Open();
        var command = new OdbcCommand();
        command.Connection = connection;
        var query = "select distinct aj4ccyy from cmsdat.glmt where aj4ccyy<2017 order by aj4ccyy desc";
        command.CommandText = query;
        var reader = command.ExecuteReader();
        while (reader.Read())
        {
            YearList.Items.Add(reader[0].ToString().Trim());
        }
        reader.Close();
        connection.Close();		
    }

    protected void ButtonSave_Click(object sender, EventArgs e)
    {
        if (GridViewSalesCAD.Visible || GridViewSalesUSD.Visible || GridViewSalesPESO.Visible)
        {
            string fileName = @"C:\Sales Report\" + User.Identity.Name + " Sales Report for plant " + plantID.ToString("D3") + " at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + ".csv";
            if (Directory.Exists(@"C:\Sales Report\"))
            {
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
            }
            else
            {
                Directory.CreateDirectory(@"C:\Sales Report\");
            }
            StreamWriter writer = new StreamWriter(fileName);
            // write sales report
            if (GridViewSalesCAD.Visible)
            {
                GenerateCSVFile(ref writer, GridViewSalesCAD);
            }
            if (GridViewSalesUSD.Visible)
            {
                GenerateCSVFile(ref writer, GridViewSalesUSD);
            }
            if (GridViewSalesPESO.Visible)
            {
                GenerateCSVFile(ref writer, GridViewSalesPESO);
            }
            if (GridViewSalesConsolidate.Visible)
            {
                GenerateCSVFile(ref writer, GridViewSalesConsolidate);
            }
            // write steel surcharge
            if (GridViewSurcharge.Visible)
            {
                GenerateCSVFile(ref writer, GridViewSurcharge);
            }
            // write piece count
            if (GridViewPieceCount.Visible)
            {
                GenerateCSVFile(ref writer, GridViewPieceCount);
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

    protected void Button_Click(object sender, EventArgs e)
    {
        Button button = (Button)sender;
        switch (button.ID)
        {
            case "ButtonMarkham":
                plantID = 1;
                GridViewSalesCAD.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Sales Report (CAD) at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewSalesUSD.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Sales Report (USD) at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewSurcharge.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Steel Surcharge (CAD) at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewPieceCount.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Piece Count at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewSalesConsolidate.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Consolidated Sales Report (CAD) at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewSalesCAD.Visible = true;
                GridViewSalesUSD.Visible = true;
                GridViewSalesPESO.Visible = false;
                GridViewSurcharge.Visible = true;
                GridViewPieceCount.Visible = true;
                GridViewSalesConsolidate.Visible = true;
                break;
            case "ButtonMichigan":
                plantID = 3;
                GridViewSalesUSD.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Sales Report (USD) at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewSurcharge.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Steel Surcharge (USD) at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewPieceCount.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Piece Count at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewSalesCAD.Visible = false;
                GridViewSalesUSD.Visible = true;
                GridViewSalesPESO.Visible = false;
                GridViewSurcharge.Visible = true;
                GridViewPieceCount.Visible = true;
                GridViewSalesConsolidate.Visible = false;
                break;
            case "ButtonColombia":
                plantID = 4;
                GridViewSalesPESO.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Sales Report (PESO) at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewSalesUSD.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Sales Report (USD) at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewSurcharge.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Steel Surcharge (PESO) at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewPieceCount.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Piece Count at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewSalesConsolidate.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Consolidated Sales Report (PESO) at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewSalesCAD.Visible = false;
                GridViewSalesUSD.Visible = true;
                GridViewSalesPESO.Visible = true;
                GridViewSurcharge.Visible = true;
                GridViewPieceCount.Visible = true;
                GridViewSalesConsolidate.Visible = true;
                break;
            case "ButtonTexas":
                plantID = 5;
                GridViewSalesUSD.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Sales Report (USD) at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewSurcharge.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Steel Surcharge (USD) at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewPieceCount.Caption = "<font size=\"9\" color=\"red\">" + button.Text + " Piece Count at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + "</font>";
                GridViewSalesCAD.Visible = false;
                GridViewSalesUSD.Visible = true;
                GridViewSalesPESO.Visible = false;
                GridViewSurcharge.Visible = true;
                GridViewSalesConsolidate.Visible = false;
                GridViewPieceCount.Visible = true;
                break;
        }
        // write sales data
        Process process = new Process();
        process.fiscalYear = Convert.ToInt32(YearList.SelectedValue.ToString()) - 2000;
        process.Run(plantID);
        double[] summaryList = new double[13];
        double[] gridList = new double[13];
        if (GridViewSalesCAD.Visible)
        {
            gridList = WriteToSalesTable(GridViewSalesCAD, process, "CA");
            for (int i = 0; i < 13; i++)
            {
                summaryList[i] += gridList[i];
            }
        }
        if (GridViewSalesUSD.Visible)
        {
            gridList = WriteToSalesTable(GridViewSalesUSD, process, "US");
            for (int i = 0; i < 13; i++)
            {
                summaryList[i] += gridList[i];
            }
        }
        if (GridViewSalesPESO.Visible)
        {
            gridList = WriteToSalesTable(GridViewSalesPESO, process, "CP");
            for (int i = 0; i < 13; i++)
            {
                summaryList[i] += gridList[i];
            }
        }
        if (GridViewSalesConsolidate.Visible)
        {
            DataTable table = new DataTable();
            table.Columns.Add(" ");
            table.Columns.Add("Period01");
            table.Columns.Add("Period02");
            table.Columns.Add("Period03");
            table.Columns.Add("Period04");
            table.Columns.Add("Period05");
            table.Columns.Add("Period06");
            table.Columns.Add("Period07");
            table.Columns.Add("Period08");
            table.Columns.Add("Period09");
            table.Columns.Add("Period10");
            table.Columns.Add("Period11");
            table.Columns.Add("Period12");
            table.Columns.Add("Total");
            List<double> actualSummary = new List<double>();
            List<double> budgetSummary = new List<double>();
            List<double> lastYearSummary = new List<double>();
            string currency;
            if (plantID == 1)
            {
                currency = "CA";
            }
            else if (plantID == 4)
            {
                currency = "CP";
            }
            else
            {
                throw new Exception("Invalid plant " + plantID);
            }
            for (int i = 0; i < 13; i++)
            {
                actualSummary.Add(0.0);
                budgetSummary.Add(0.0);
                lastYearSummary.Add(0.0);
            }
            DataRow rowActual = table.NewRow();
            rowActual[" "] = "Actual:";
            DataRow rowBudget = table.NewRow();
            rowBudget[" "] = "Budget:";
            DataRow rowLastYear = table.NewRow();
            rowLastYear[" "] = "Last Year:";
            foreach (Customer cust in process.plant.custList)
            {
                for (int i = 0; i <= 12; i++)
                {
                    actualSummary[i] += cust.actualList[i].GetAmount(currency);
                    budgetSummary[i] += cust.budgetList[i].GetAmount(currency);
                    lastYearSummary[i] += cust.actualListLastYear[i].GetAmount(currency);
                }
            }
            double actualTotal = 0.0;
            double budgetTotal = 0.0;
            double lastYearTotal = 0.0;
            for (int i = 1; i <= 12; i++)
            {
                rowActual["Period" + i.ToString("D2")] = actualSummary[i].ToString("C2");
                actualTotal += actualSummary[i];
                rowBudget["Period" + i.ToString("D2")] = budgetSummary[i].ToString("C2");
                budgetTotal += budgetSummary[i];
                rowLastYear["Period" + i.ToString("D2")] = lastYearSummary[i].ToString("C2");
                lastYearTotal += lastYearSummary[i];
            }
            rowActual["Total"] = actualTotal.ToString("C2");
            rowBudget["Total"] = budgetTotal.ToString("C2");
            rowLastYear["Total"] = lastYearTotal.ToString("C2");
            table.Rows.Add(rowActual);
            table.Rows.Add(rowBudget);
            table.Rows.Add(rowLastYear);
            // write to grid view
            table.AcceptChanges();
            GridViewSalesConsolidate.DataSource = table;
            GridViewSalesConsolidate.DataBind();
            // adjust style
            foreach (GridViewRow row in GridViewSalesConsolidate.Rows)
            {
                Style style = new Style();
                style.ForeColor = Color.Gray;
                style.Font.Size = 9;
                row.ApplyStyle(style);
            }
        }
        // write steel surcharge
        WriteSteelSurchargeTable(GridViewSurcharge, process, summaryList);
        // write piece count
        WritePieceCountTable(GridViewPieceCount, process);
    }

    private void WritePieceCountTable(GridView gridView, Process process)
    {
        DataTable table = new DataTable();
        table.Columns.Add(" ");
        table.Columns.Add("Period01");
        table.Columns.Add("Period02");
        table.Columns.Add("Period03");
        table.Columns.Add("Period04");
        table.Columns.Add("Period05");
        table.Columns.Add("Period06");
        table.Columns.Add("Period07");
        table.Columns.Add("Period08");
        table.Columns.Add("Period09");
        table.Columns.Add("Period10");
        table.Columns.Add("Period11");
        table.Columns.Add("Period12");
        table.Columns.Add("Total");
        // add solid
        DataRow solidRow = table.NewRow();
        solidRow[0] = "Solid:";
        int total = 0;
        for (int i = 1; i <= 12; i++)
        {
            solidRow[i] = process.plant.solidList[i].ToString();
            total += process.plant.solidList[i];
        }
        solidRow["Total"] = total.ToString();
        // add hollow
        DataRow hollowRow = table.NewRow();
        hollowRow[0] = "Hollow:";
        total = 0;
        for (int i = 1; i <= 12; i++)
        {
            hollowRow[i] = process.plant.hollowList[i].ToString();
            total += process.plant.hollowList[i];
        }
        hollowRow["Total"] = total.ToString();
        // add ncr
        DataRow ncrRow = table.NewRow();
        ncrRow[0] = "Ncr:";
        total = 0;
        for (int i = 1; i <= 12; i++)
        {
            ncrRow[i] = process.plant.ncrList[i].ToString();
            total += process.plant.ncrList[i];
        }
        ncrRow["Total"] = total.ToString();
        // write to grid view
        table.Rows.Add(solidRow);
        table.Rows.Add(hollowRow);
        table.Rows.Add(ncrRow);
        table.AcceptChanges();
        gridView.DataSource = table;
        gridView.DataBind();
    }

    private void WriteSteelSurchargeTable(GridView gridView, Process process, double[] summaryList)
    {
        DataTable table = new DataTable();
        table.Columns.Add(" ");
        table.Columns.Add("Period01");
        table.Columns.Add("Period02");
        table.Columns.Add("Period03");
        table.Columns.Add("Period04");
        table.Columns.Add("Period05");
        table.Columns.Add("Period06");
        table.Columns.Add("Period07");
        table.Columns.Add("Period08");
        table.Columns.Add("Period09");
        table.Columns.Add("Period10");
        table.Columns.Add("Period11");
        table.Columns.Add("Period12");
        table.Columns.Add("Total");
        // add surcharge amount
        DataRow amountRow = table.NewRow();
        amountRow[0] = "Surcharge:";
        double total = 0.0;
        for (int i = 1; i <= 12; i++)
        {
            amountRow[i] = process.plant.surchargeList[i].GetAmount(process.plant.currency).ToString("C0");
            total += process.plant.surchargeList[i].GetAmount(process.plant.currency);
        }
        amountRow["Total"] = total.ToString("C0");
        // add ratio
        DataRow ratioRow = table.NewRow();
        ratioRow[0] = "Percentage:";
        for (int i = 1; i <= 12; i++)
        {
            if (summaryList[i - 1] < 0.001)
            {
                ratioRow[i] = "0%";
            }
            else
            {
                ratioRow[i] = (process.plant.surchargeList[i].GetAmount(process.plant.currency) / summaryList[i - 1]).ToString("P2");
            }
        }
        if (summaryList[12] < 0.001)
        {
            ratioRow["Total"] = "0%";
        }
        else
        {
            ratioRow["Total"] = (total / summaryList[12]).ToString("P2");
        }
        // write to grid view
        table.Rows.Add(amountRow);
        table.Rows.Add(ratioRow);
        table.AcceptChanges();
        gridView.DataSource = table;
        gridView.DataBind();
    }

    private double[] WriteToSalesTable(GridView gridView, Process process, string currency)
    {
        DataTable table = new DataTable();
        table.Columns.Add("Cust#");
        table.Columns.Add("Cust Name");
        table.Columns.Add("Territory");
        table.Columns.Add("Currency");
        table.Columns.Add(" ");
        table.Columns.Add("Period01");
        table.Columns.Add("Period02");
        table.Columns.Add("Period03");
        table.Columns.Add("Period04");
        table.Columns.Add("Period05");
        table.Columns.Add("Period06");
        table.Columns.Add("Period07");
        table.Columns.Add("Period08");
        table.Columns.Add("Period09");
        table.Columns.Add("Period10");
        table.Columns.Add("Period11");
        table.Columns.Add("Period12");
        table.Columns.Add("Total");
        List<ExcoMoney> actualSummary = new List<ExcoMoney>();
        List<ExcoMoney> budgetSummary = new List<ExcoMoney>();
        List<ExcoMoney> lastYearSummary = new List<ExcoMoney>();
        for (int i = 0; i < 13; i++)
        {
            actualSummary.Add(new ExcoMoney());
            budgetSummary.Add(new ExcoMoney());
            lastYearSummary.Add(new ExcoMoney());
        }
        // add sales data
        foreach (Customer cust in process.plant.custList)
        {
            if (cust.excoCustomer.Currency.Contains(currency) && (!cust.actualTotal.IsZero() || !cust.budgetTotal.IsZero()))
            {
                DataRow rowActual = table.NewRow();
                rowActual["Cust#"] = cust.excoCustomer.BillToID;
                rowActual["Cust Name"] = cust.excoCustomer.Name;
                rowActual["Territory"] = cust.excoCustomer.Territory;
                rowActual["Currency"] = cust.excoCustomer.Currency;
                rowActual[" "] = "Actual:";
                DataRow rowBudget = table.NewRow();
                rowBudget[" "] = "Budget:";
                DataRow rowLastYear = table.NewRow();
                rowLastYear[" "] = "Last Year:";
                for (int i = 1; i <= 12; i++)
                {
                    double actual = cust.actualList[i].GetAmount(currency);
                    double budget = cust.budgetList[i].GetAmount(currency);
                    double lastYear = cust.actualListLastYear[i].GetAmount(currency);
                    string index = "Period" + i.ToString("D2");
                    rowActual[index] = actual.ToString("C2");
                    actualSummary[i - 1] += cust.actualList[i];
                    rowBudget[index] = budget.ToString("C2");
                    budgetSummary[i - 1] += cust.budgetList[i];
                    rowLastYear[index] = lastYear.ToString("C2");
                    lastYearSummary[i - 1] += cust.actualListLastYear[i];
                }
                rowActual["Total"] = cust.actualTotal.GetAmount(currency).ToString("C2");
                actualSummary[12] += cust.actualTotal;
                rowBudget["Total"] = cust.budgetTotal.GetAmount(currency).ToString("C2");
                budgetSummary[12] += cust.budgetTotal;
                rowLastYear["Total"] = cust.actualTotalLastYear.GetAmount(currency).ToString("C2");
                lastYearSummary[12] += cust.actualTotalLastYear;
                table.Rows.Add(rowActual);
                table.Rows.Add(rowBudget);
                table.Rows.Add(rowLastYear);
            }
        }
        // add empty rows
        table.Rows.Add(table.NewRow());
        table.Rows.Add(table.NewRow());
        // add summary
        double[] summaryList = { 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 };
        DataRow summaryActual = table.NewRow();
        DataRow summaryBudget = table.NewRow();
        DataRow summaryLastYear = table.NewRow();
        summaryActual["Currency"] = "TOTAL (" + currency + "):";
        summaryActual[" "] = "Actual:";
        summaryBudget[" "] = "Budget:";
        summaryLastYear[" "] = "Last Year:";
        for (int i = 0; i < 13; i++)
        {
            summaryActual[i + 5] = actualSummary[i].GetAmount(currency).ToString("C2");
            summaryBudget[i + 5] = budgetSummary[i].GetAmount(currency).ToString("C2");
            summaryLastYear[i + 5] = lastYearSummary[i].GetAmount(currency).ToString("C2");
            if ((1 == plantID && currency.Contains("CA")) || (4 == plantID && currency.Contains("CP")) || 3 == plantID || 5 == plantID)
            {
                summaryList[i] += actualSummary[i].GetAmount(currency);
            }
        }   
        table.Rows.Add(summaryActual);
        table.Rows.Add(summaryBudget);
        table.Rows.Add(summaryLastYear);
        // add currency conversion
        if (1 == plantID && currency.Contains("US"))
        {
            // add empty rows
            table.Rows.Add(table.NewRow());
            table.Rows.Add(table.NewRow());
            // add conversion rate
            DataRow rateRow = table.NewRow();
            DataRow actualRow = table.NewRow();
            DataRow budgetRow = table.NewRow();
            DataRow lastYearRow = table.NewRow();
            rateRow["Currency"] = "Convert to CAD";
            rateRow[" "] = "Rate:";
            actualRow[" "] = "Actual:";
            budgetRow[" "] = "Budget:";
            lastYearRow[" "] = "Last Year:";
            for (int i = 1; i <= 12; i++)
            {
                rateRow["Period" + i.ToString("D2")] = ExcoExRate.GetToCADRate(new ExcoCalendar(14, i, true, 1), "US").ToString("F2");
                actualRow["Period" + i.ToString("D2")] = actualSummary[i - 1].GetAmount("CA").ToString("C2");
                budgetRow["Period" + i.ToString("D2")] = budgetSummary[i - 1].GetAmount("CA").ToString("C2");
                lastYearRow["Period" + i.ToString("D2")] = lastYearSummary[i - 1].GetAmount("CA").ToString("C2");
                summaryList[i - 1] += actualSummary[i - 1].GetAmount("CA");
            }
            actualRow["Total"] = actualSummary[12].GetAmount("CA").ToString("C2");
            summaryList[12] += actualSummary[12].GetAmount("CA");
            budgetRow["Total"] = budgetSummary[12].GetAmount("CA").ToString("C2");
            lastYearRow["Total"] = lastYearSummary[12].GetAmount("CA").ToString("C2");
            table.Rows.Add(rateRow);
            table.Rows.Add(actualRow);
            table.Rows.Add(budgetRow);
            table.Rows.Add(lastYearRow);
        }
        if (4 == plantID && currency.Contains("US"))
        {
            // add empty rows
            table.Rows.Add(table.NewRow());
            table.Rows.Add(table.NewRow());
            // add conversion rate
            DataRow rateRow = table.NewRow();
            DataRow actualRow = table.NewRow();
            DataRow budgetRow = table.NewRow();
            DataRow lastYearRow = table.NewRow();
            rateRow["Currency"] = "Convert to PESO";
            rateRow[" "] = "Rate:";
            actualRow[" "] = "Actual:";
            budgetRow[" "] = "Budget:";
            for (int i = 1; i <= 12; i++)
            {
                if (i <= 3)
                {
                    rateRow["Period" + i.ToString("D2")] = ExcoExRate.GetToPESORate(new ExcoCalendar(13, i + 9, true, 4), "US").ToString("F2");
                }
                else
                {
                    rateRow["Period" + i.ToString("D2")] = ExcoExRate.GetToPESORate(new ExcoCalendar(14, i - 3, true, 4), "US").ToString("F2");
                }
                actualRow["Period" + i.ToString("D2")] = actualSummary[i - 1].GetAmount("CP").ToString("C2");
                summaryList[i - 1] += actualSummary[i - 1].GetAmount("CP");
                budgetRow["Period" + i.ToString("D2")] = budgetSummary[i - 1].GetAmount("CP").ToString("C2");
                lastYearRow["Period" + i.ToString("D2")] = lastYearSummary[i - 1].GetAmount("CP").ToString("C2");
            }
            actualRow["Total"] = actualSummary[12].GetAmount("CP").ToString("C2");
            summaryList[12] += actualSummary[12].GetAmount("CP");
            budgetRow["Total"] = budgetSummary[12].GetAmount("CP").ToString("C2");
            lastYearRow["Total"] = lastYearSummary[12].GetAmount("CP").ToString("C2");
            table.Rows.Add(rateRow);
            table.Rows.Add(actualRow);
            table.Rows.Add(budgetRow);
            table.Rows.Add(lastYearRow);
        }
        // write to grid view
        table.AcceptChanges();
        gridView.DataSource = table;
        gridView.DataBind();
        // adjust style
        foreach (GridViewRow row in gridView.Rows)
        {
            Style style = new Style();
            style.ForeColor = Color.Gray;
            style.Font.Size = 9;
            // actual row
            if (row.Cells[4].Text.Contains("Actual") || row.Cells[4].Text.Contains("Rate"))
            {
                row.ForeColor = Color.Black;
                row.Font.Bold = true;
                row.Font.Size = 12;
                row.Cells[4].ForeColor = Color.Gray;
                row.Cells[4].Font.Size = 9;
            }
            // budget/last year row
            else
            {
                row.ApplyStyle(style);
            }
        }
        return summaryList;
    }
}