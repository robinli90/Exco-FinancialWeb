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
using IncomeStatementReport;
using System.Globalization;

public partial class income_statement_report : System.Web.UI.Page
{

    private bool Implement_Changes = true;


    private int fiscalYear;
    private int fiscalMonth;
    private string currency;
    private int[,] blockposition = new int[20, 2];       //blockposition[i,0] block row start; blockposition[i,1] block row stop;
    private int[] totalposition = new int[12];            //total line of: cost of goods, gross marging, total expenses, total net income
    private double[] provisionalTax = new double[13];            
    private double[,] allsummary = new double[12, 12];
    private string tempdata;
    private double tempp;
    private NumberStyles styles = NumberStyles.Currency;
    private double ytdActual = 0.0;
    private double ytdBudget = 0.0;

    // rows for consolidation
    private int provisionRow = 0;
    private int factorySATravelRow = 0;
    private int officeSATravelRow = 0;
    private int grandTotalRow = 0;
    private int directLabourShop = 0;
    private int directLabourHeat = 0;
    private int directLabourCAM = 0;
    private int factoryCAD = 0;
    private int factoryIndirectLabour = 0;
    private int factorySupervisor = 0;
    private int factoryEmployeeBenefits = 0;
    private int factoryGroupInsurance = 0;
    private int factoryWorkersCompensation = 0;
    private int factoryVacationPay = 0;
    private int factoryStatHoliday = 0;
    private int officeSalaries = 0;
    private int salesSalaries = 0;
    private int salesEmployeeBenefits = 0;
    private int salesGroupInsurance = 0;
    private int salesWorkersCompensation = 0;
    private int salesVacationPay = 0;
    private int salesStatHoliday = 0;
    private int officeEmployeeBenefits = 0;
    private int officeGroupInsurance = 0;
    private int officePayrollTaxes = 0;
    private int officeWorkersCompensation = 0;
    private int officeVacationPay = 0;
    private int officeStatHoliday = 0;
    private int factoryRent = 0;
    private int officeRent = 0;
    private int factoryDepreBuilding = 0;
    private int factoryDepreBuildingImprov = 0;
    private int factoryReqMainBuilding = 0;
    private int factoryPropertyTax = 0;
    private int factoryDepreMachine = 0;
    private int factoryDepreFurniture = 0;
    private int factoryDepreSoftware = 0;
    private int factoryToolAmort = 0;
    private int officeDepreFurniture = 0;
    private int officeSoftwareAmort = 0;
    private int salesDepreAuto = 0;
    private int factoryToolExpense = 0;
    private int factoryCADCAMSupplies = 0;
    private int factoryShopSupplies = 0;
    private int factoryHeatTreatSupplies = 0;
    private int factoryShippingSupplies = 0;
    private int officeSupplies = 0;
    private int factoryTravel = 0;
    private int factoryMeals = 0;
    private int factoryAirFare = 0;
    private int salesGolf = 0;
    private int salesMeals = 0;
    private int salesTravel = 0;
    private int salesAirFare = 0;
    private int officeMeals = 0;
    private int officeTravel = 0;
    private int officeAirFare = 0;
    private int row = 0;
    private int switchpositivenegative = -1;

    int[] total_line_row_number = new int[4];
    int[] line_subgroup = new int[4];


    protected void Page_Load(object sender, EventArgs e)
    {
        //ExcoExRate.GetExchangeRatesFromFile();

        SelectConsolidate.Visible = true;

        if (true == (bool)Session["_hasIncomeStatementReportTexas"])
        {
            SelectTexas.Visible = true;
        }
        if (true == (bool)Session["_hasIncomeStatementReportColombia"])
        {
            SelectColombia.Visible = true;
        }
        if (true == (bool)Session["_hasIncomeStatementReportMichigan"])
        {
            SelectMichigan.Visible = true;
        }
        if (true == (bool)Session["_hasIncomeStatementReportMarkham"])
        {
            SelectMarkham.Visible = true;
        }

        // load all fiscal years
        var connection = new OdbcConnection(ConfigurationManager.ConnectionStrings["cms1"].ConnectionString);
        connection.Open();
        var command = new OdbcCommand();
        command.Connection = connection;
        var query = "select distinct aj4ccyy from cmsdat.glmt where aj4ccyy<2019 order by aj4ccyy desc";
        command.CommandText = query;
        var reader = command.ExecuteReader();
        while (reader.Read())
        {
            YearList.Items.Add(reader[0].ToString().Trim());
        }
        reader.Close();
        connection.Close();		
    }

    protected void ButtonGenerate_Click(object sender, EventArgs e)
    {
        fiscalYear = Convert.ToInt32(YearList.SelectedValue.ToString())-2000;
        fiscalMonth = Convert.ToInt32(PeriodList.SelectedValue.ToString());

        if (SelectMarkham.Checked)
        {
            MarkhamReport();
        }
        else if (SelectMichigan.Checked)
        {
            MichiganReport();
        }
        else if (SelectTexas.Checked)
        {
            TexasReport();
        }
        else if (SelectColombia.Checked)
        {
            ColombiaReport();
        }
        else if (SelectConsolidate.Checked)
        {
            ConsolidateReport();
        }
    }

    private void MarkhamReport()
    {
        Session["_plantID"] = "001";
        Session["_plantName"] = "Markham";
        currency = "CA";
        GridView gridView = GridViewIncomeStatementReport;
        ExcoCalendar calendar = new ExcoCalendar(fiscalYear, fiscalMonth, true, 1);
        DateTime datename = new DateTime(fiscalYear+2000,calendar.GetCalendarMonth(),1);
        string monthname = datename.ToString("MMM");
        string yearname;
        if (monthname == "Oct" || monthname == "Nov" || monthname == "Dec")
        {
            yearname = (fiscalYear + 1999).ToString();
        }
        else
        {
            yearname = (fiscalYear + 2000).ToString();
        }
        gridView.Caption = "<font size=\"9\" color=\"red\">Markham Income Statement Report (CAD) Period " + fiscalMonth + " - " + (fiscalYear + 2000) + " Ending " + monthname + " " + DateTime.DaysInMonth(fiscalYear + 2000, calendar.GetCalendarMonth()) + ", " + yearname + "</font>";
        PanelIncomeStatementReport.Visible = true;
        gridView.Visible = true;
        //write table data
        DataTable table = new DataTable();
        //header column
        WriteHeaderByPlant(table);
        //call .dll to process all the table data
        Process process = new Process(fiscalYear, fiscalMonth);

        //Block 1 : Sales                     
        WriteAmount("SALES", table, 1, process.ss.groupList, 0, currency);

        //Block 2 : Cost of Steels
        WriteAmount("COST OF STEELS", table, 1, process.cs.groupList, 1, currency);

        //Block 3 : Cost of Steels
        WriteAmount("DIRECT LABOUR", table, 1, process.dl.groupList, 2, currency);

        //Block 4 : Factory Overhead
        WriteAmount("FACTORY OVERHEAD", table, 1, process.fo.groupList, 3, currency);

        //Total cost of goods
        WriteTotalLine(table, 0);

        //Gross Margin
        WriteTotalLine(table, 1);

        //Block 5 : Delivery and selling
        WriteAmount("DELIVERY AND SELLING", table, 1, process.ds.groupList, 4, currency);

        //Block 6 : General and administration
        WriteAmount("GENERAL AND ADMINISTRATION", table, 1, process.ga.groupList, 5, currency);

        //Block 7 : Other expenses and income
        WriteAmount("OTHER EXPENSES AND INCOME", table, 1, process.oe.groupList, 6, currency);

        //Total expenses
        WriteTotalLine(table, 2);

        //net income before provision tax
        WriteTotalLine(table, 3);

        //Block 8 : Provisional Taxes
        WriteAmount("", table, 1, process.pt.groupList, 7, currency);

        //Total net income
        WriteTotalLine(table, 4);

        table.AcceptChanges();
        gridView.DataSource = table;
        gridView.DataBind();

        //style adjust
        TableStyleAdjustment(table, gridView); 
    }

    protected void ButtonSave_Click(object sender, EventArgs e)
    {
        // write to excel
        string fileName = @"C:\Income Statement Report\" + User.Identity.Name + " Income Statement Report for " + Session["_plantName"].ToString() + " at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + ".csv";

        if (Directory.Exists(@"C:\Income Statement Report\"))
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
        if (GridViewIncomeStatementReport.Visible)
        {
            GenerateCSVFile(ref writer, GridViewIncomeStatementReport);
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

    protected void ConsolidateReport()
    {
        GridView gridView = GridViewIncomeStatementReport;
        ExcoCalendar calendar = new ExcoCalendar(fiscalYear, fiscalMonth, true, 1);
        DateTime datename = new DateTime(fiscalYear + 2000, calendar.GetCalendarMonth(), 1);
        string monthname = datename.ToString("MMM");
        string yearname;
        if (monthname == "Oct" || monthname == "Nov" || monthname == "Dec")
        {
            yearname = (fiscalYear + 1999).ToString();
        }
        else
        {
            yearname = (fiscalYear + 2000).ToString();
        }
        gridView.Caption = "<font size=\"9\" color=\"red\">Consolidated Income Statement Report (CAD) Period " + fiscalMonth + " - " + (fiscalYear + 2000) + " Ending " + monthname + " " + DateTime.DaysInMonth(fiscalYear + 2000, calendar.GetCalendarMonth()) + ", " + yearname + "</font>";
        PanelIncomeStatementReport.Visible = true;
        gridView.Visible = true;
        //write table data
        DataTable table = new DataTable();
        table.Columns.Add("Name");                          
        table.Columns.Add("Markham");               
        table.Columns.Add(" ");                         
        table.Columns.Add("Michigan");            
        table.Columns.Add("  ");                      
        table.Columns.Add("Texas");               
        table.Columns.Add("   ");                          
        table.Columns.Add("Colombia");                  
        table.Columns.Add("    ");                          
        table.Columns.Add("Total");                 
        table.Columns.Add("     ");
        table.Columns.Add("YTD Total");

        DataRow exchange = table.NewRow();
        exchange[0] = "Exchange Rate:";
        exchange[1] = 1.0;
        //ExcoCalendar calendar = new ExcoCalendar(fiscalYear, fiscalMonth, true, 1);
        exchange[3] = ExcoExRate.GetToCADRate(calendar.GetNextCalendarMonth(), "US");
        exchange[5] = ExcoExRate.GetToCADRate(calendar.GetNextCalendarMonth(), "US");
        exchange[7] = ExcoExRate.GetToCADRate(calendar.GetNextCalendarMonth(), "CP");
        table.Rows.InsertAt(exchange,0);

        DataRow emprow = table.NewRow();
        //table.Rows.Add(emprow);

        Process process = new Process(fiscalYear, fiscalMonth, false);
        //Block 1 : Sales 
        WriteAmountConsolidate("SALES", table, process.ss.groupList, 0, calendar);

        //Block 2 : Cost of Steels
        WriteAmountConsolidate("COST OF STEELS", table, process.cs.groupList, 1, calendar);
       
        //Block 3 : Cost of Steels
        WriteAmountConsolidate("DIRECT LABOUR", table, process.dl.groupList, 2, calendar);

        //Block 4 : Factory Overhead
        WriteAmountConsolidate("FACTORY OVERHEAD", table, process.fo.groupList, 3, calendar);

        //Total cost of goods
        WriteTotalLineConsolidate(table, 0);

        //Gross Margin
        WriteTotalLineConsolidate(table, 1);

        //Block 5 : Delivery and selling
        WriteAmountConsolidate("DELIVERY AND SELLING", table, process.ds.groupList, 4, calendar);

        //Block 6 : General and administration
        WriteAmountConsolidate("GENERAL AND ADMINISTRATION", table, process.ga.groupList, 5, calendar);

        //Block 7 : Other expenses and income
        WriteAmountConsolidate("OTHER EXPENSES AND INCOME", table, process.oe.groupList,6, calendar);

        //Total expenses
        WriteTotalLineConsolidate(table, 2);

        //Total net income before income tax
        WriteTotalLineConsolidate(table, 3);

        //Block 8: provision
        WriteAmountConsolidate("", table, process.pt.groupList, 7, calendar);

        //Total net income
        WriteTotalLineConsolidate(table, 4);

        /*
        #region block adjustment
        DataRow adjust = table.NewRow();
        adjust[0] = "ADJUSTMENTS:";
        table.Rows.Add(adjust);

        blockposition[7, 0] = table.Rows.Count - 1;

        DataRow adjust1 = table.NewRow();
        adjust1[0] = "ADD: INCOME TAX PROVISION:";
        adjust1[1] = table.Rows[provisionRow][1];
        adjust1[3] = table.Rows[provisionRow][3];
        adjust1[5] = table.Rows[provisionRow][5];
        adjust1[7] = table.Rows[provisionRow][7];
        adjust1[9] = table.Rows[provisionRow][9];

        tempdata = table.Rows[provisionRow][1].ToString();
        tempp = double.Parse(tempdata, styles);
        adjust1[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = table.Rows[provisionRow][3].ToString();
        tempp = double.Parse(tempdata, styles);
        adjust1[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = table.Rows[provisionRow][5].ToString();
        tempp = double.Parse(tempdata, styles);
        adjust1[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = table.Rows[provisionRow][7].ToString();
        tempp = double.Parse(tempdata, styles);
        adjust1[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = table.Rows[provisionRow][9].ToString();
        tempp = double.Parse(tempdata, styles);
        adjust1[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(adjust1);

        DataRow adjust2 = table.NewRow();
        adjust2[0] = "RECLASS: SOUTH AMERICA TRAVEL";

        tempdata = table.Rows[factorySATravelRow][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[officeSATravelRow][1].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust2[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factorySATravelRow][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[officeSATravelRow][3].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust2[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factorySATravelRow][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[officeSATravelRow][5].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust2[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factorySATravelRow][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[officeSATravelRow][7].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust2[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factorySATravelRow][9].ToString();
        //tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[officeSATravelRow][9].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust2[9] = NegativeCurrencyFormat(tempp);

        tempdata = adjust2[1].ToString();
        tempp = double.Parse(tempdata, styles);
        adjust2[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = adjust2[3].ToString();
        tempp = double.Parse(tempdata, styles);
        adjust2[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = adjust2[5].ToString();
        tempp = double.Parse(tempdata, styles);
        adjust2[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = adjust2[7].ToString();
        tempp = double.Parse(tempdata, styles);
        adjust2[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = adjust2[9].ToString();
        tempp = double.Parse(tempdata, styles);
        adjust2[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(adjust2);

        DataRow adjust3 = table.NewRow();
        adjust3[0] = "TOTAL ADJUSTMENTS:";

        row = table.Rows.Count;

        tempdata = table.Rows[row - 2][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust3[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust3[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust3[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust3[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][9].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust3[9] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust3[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = table.Rows[row - 2][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust3[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = table.Rows[row - 2][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust3[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = table.Rows[row - 2][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust3[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = table.Rows[row - 2][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][9].ToString();
        tempp += double.Parse(tempdata, styles);
        adjust3[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(adjust3);

        blockposition[7, 1] = table.Rows.Count - 1;

        DataRow emp = table.NewRow();
        table.Rows.Add(emp);
         * 
         * 
        DataRow adjust4 = table.NewRow();
        //adjust4[0] = "ADJUSTED NET INCOME BEFORE INC. TAX:";
        adjust4[0] = "INCOME BEFORE PROVOSION TAX:";

        tempdata = table.Rows[totalposition[3]][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[7, 1]][1].ToString();
        tempp -= double.Parse(tempdata, styles);
        adjust4[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[totalposition[3]][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[7, 1]][3].ToString();
        tempp -= double.Parse(tempdata, styles);
        adjust4[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[totalposition[3]][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[7, 1]][5].ToString();
        tempp -= double.Parse(tempdata, styles);
        adjust4[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[totalposition[3]][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[7, 1]][7].ToString();
        tempp -= double.Parse(tempdata, styles);
        adjust4[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[totalposition[3]][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[7, 1]][9].ToString();
        tempp -= double.Parse(tempdata, styles);
        adjust4[9] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[totalposition[3]][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[7, 1]][1].ToString();
        tempp -= double.Parse(tempdata, styles);
        adjust4[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = table.Rows[totalposition[3]][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[7, 1]][3].ToString();
        tempp -= double.Parse(tempdata, styles);
        adjust4[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = table.Rows[totalposition[3]][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[7, 1]][5].ToString();
        tempp -= double.Parse(tempdata, styles);
        adjust4[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = table.Rows[totalposition[3]][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[7, 1]][7].ToString();
        tempp -= double.Parse(tempdata, styles);
        adjust4[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = table.Rows[totalposition[3]][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[7, 1]][9].ToString();
        tempp -= double.Parse(tempdata, styles);
        adjust4[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(adjust4);

        totalposition[4] = table.Rows.Count - 1;


        */

        DataRow emp1 = table.NewRow();
        table.Rows.Add(emp1);
        DataRow emp2 = table.NewRow();
        table.Rows.Add(emp2);
        DataRow emp3 = table.NewRow();
        table.Rows.Add(emp3);


        #region supply title
        DataRow supply = table.NewRow();
        supply[0] = "SUPPLEMENTAL INFO:";
        table.Rows.Add(supply);
        totalposition[5] = table.Rows.Count - 1;
        DataRow emp4 = table.NewRow();
        table.Rows.Add(emp4);
        #endregion

        #region SHOP LABOR
        DataRow shop = table.NewRow();
        shop[0] = "SHOP LABOR";
        table.Rows.Add(shop);

        blockposition[8, 0] = table.Rows.Count - 1;

        DataRow shop1 = table.NewRow();
        shop1[0] = "WAGES";
        tempdata = table.Rows[directLabourShop][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[directLabourHeat][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[directLabourCAM][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryCAD][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryIndirectLabour][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factorySupervisor][1].ToString();
        tempp += double.Parse(tempdata, styles);
        shop1[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[directLabourShop][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[directLabourHeat][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[directLabourCAM][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryCAD][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryIndirectLabour][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factorySupervisor][3].ToString();
        tempp += double.Parse(tempdata, styles);
        shop1[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[directLabourShop][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[directLabourHeat][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[directLabourCAM][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryCAD][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryIndirectLabour][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factorySupervisor][5].ToString();
        tempp += double.Parse(tempdata, styles);
        shop1[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[directLabourShop][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[directLabourHeat][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[directLabourCAM][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryCAD][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryIndirectLabour][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factorySupervisor][7].ToString();
        tempp += double.Parse(tempdata, styles);
        shop1[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[directLabourShop][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[directLabourHeat][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[directLabourCAM][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryCAD][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryIndirectLabour][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factorySupervisor][9].ToString();
        tempp += double.Parse(tempdata, styles);
        shop1[9] = NegativeCurrencyFormat(tempp);

        tempdata = shop1[1].ToString();
        tempp = double.Parse(tempdata, styles);
        shop1[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = shop1[3].ToString();
        tempp = double.Parse(tempdata, styles);
        shop1[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = shop1[5].ToString();
        tempp = double.Parse(tempdata, styles);
        shop1[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = shop1[7].ToString();
        tempp = double.Parse(tempdata, styles);
        shop1[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = shop1[9].ToString();
        tempp = double.Parse(tempdata, styles);
        shop1[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(shop1);

        DataRow shop2 = table.NewRow();
        shop2[0] = "BENEFITS (% IS TO WAGES)";

        tempdata = table.Rows[factoryEmployeeBenefits][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryGroupInsurance][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryStatHoliday][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryVacationPay][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryWorkersCompensation][1].ToString();
        tempp += double.Parse(tempdata, styles);
        shop2[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryEmployeeBenefits][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryGroupInsurance][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryStatHoliday][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryVacationPay][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryWorkersCompensation][3].ToString();
        tempp += double.Parse(tempdata, styles);
        shop2[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryEmployeeBenefits][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryGroupInsurance][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryStatHoliday][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryVacationPay][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryWorkersCompensation][5].ToString();
        tempp += double.Parse(tempdata, styles);
        shop2[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryEmployeeBenefits][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryGroupInsurance][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryStatHoliday][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryVacationPay][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryWorkersCompensation][7].ToString();
        tempp += double.Parse(tempdata, styles);
        shop2[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryEmployeeBenefits][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryGroupInsurance][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryStatHoliday][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryVacationPay][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryWorkersCompensation][9].ToString();
        tempp += double.Parse(tempdata, styles);
        shop2[9] = NegativeCurrencyFormat(tempp);

        row = table.Rows.Count;

        tempdata = shop2[1].ToString(); 
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp /= double.Parse(tempdata, styles);
        shop2[2] = tempp.ToString("P2");

        tempdata = shop2[3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp /= double.Parse(tempdata, styles);
        shop2[4] = tempp.ToString("P2");

        tempdata = shop2[5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp /= double.Parse(tempdata, styles);
        shop2[6] = tempp.ToString("P2");

        tempdata = shop2[7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp /= double.Parse(tempdata, styles);
        shop2[8] = tempp.ToString("P2");

        tempdata = shop2[9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][9].ToString();
        tempp /= double.Parse(tempdata, styles);
        shop2[10] = tempp.ToString("P2");

        table.Rows.Add(shop2);

        DataRow shop3 = table.NewRow();
        shop3[0] = "TOTAL SHOP LABOR COST:";

        row = table.Rows.Count;

        tempdata = table.Rows[row - 2][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp += double.Parse(tempdata, styles);
        shop3[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp += double.Parse(tempdata, styles);
        shop3[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp += double.Parse(tempdata, styles);
        shop3[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp += double.Parse(tempdata, styles);
        shop3[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][9].ToString();
        tempp += double.Parse(tempdata, styles);
        shop3[9] = NegativeCurrencyFormat(tempp);

        tempdata = shop3[1].ToString();
        tempp = double.Parse(tempdata, styles);
        shop3[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = shop3[3].ToString();
        tempp = double.Parse(tempdata, styles);
        shop3[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = shop3[5].ToString();
        tempp = double.Parse(tempdata, styles);
        shop3[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = shop3[7].ToString();
        tempp = double.Parse(tempdata, styles);
        shop3[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = shop3[9].ToString();
        tempp = double.Parse(tempdata, styles);
        shop3[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(shop3);

        blockposition[8, 1] = table.Rows.Count - 1;

        DataRow emp5 = table.NewRow();
        table.Rows.Add(emp5);

        #endregion

        #region S,G & A LABOR
        DataRow sga = table.NewRow();
        sga[0] = "S,G & A LABOR";
        table.Rows.Add(sga);

        blockposition[9, 0] = table.Rows.Count - 1;

        DataRow sga1 = table.NewRow();
        sga1[0] = "WAGES";
        tempdata = table.Rows[officeSalaries][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[salesSalaries][1].ToString();
        tempp += double.Parse(tempdata, styles);
        sga1[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeSalaries][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[salesSalaries][3].ToString();
        tempp += double.Parse(tempdata, styles);
        sga1[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeSalaries][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[salesSalaries][5].ToString();
        tempp += double.Parse(tempdata, styles);
        sga1[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeSalaries][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[salesSalaries][7].ToString();
        tempp += double.Parse(tempdata, styles);
        sga1[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeSalaries][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[salesSalaries][9].ToString();
        tempp += double.Parse(tempdata, styles);
        sga1[9] = NegativeCurrencyFormat(tempp);

        tempdata = sga1[1].ToString();
        tempp = double.Parse(tempdata, styles);
        sga1[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = sga1[3].ToString();
        tempp = double.Parse(tempdata, styles);
        sga1[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = sga1[5].ToString();
        tempp = double.Parse(tempdata, styles);
        sga1[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = sga1[7].ToString();
        tempp = double.Parse(tempdata, styles);
        sga1[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = sga1[9].ToString();
        tempp = double.Parse(tempdata, styles);
        sga1[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(sga1);

        DataRow sga2 = table.NewRow();
        sga2[0] = "BENEFITS (% IS TO WAGES)";
        tempdata = table.Rows[salesEmployeeBenefits][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[salesGroupInsurance][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesStatHoliday][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesVacationPay][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesWorkersCompensation][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeEmployeeBenefits][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeGroupInsurance][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officePayrollTaxes][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeVacationPay][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeWorkersCompensation][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeStatHoliday][1].ToString();
        tempp += double.Parse(tempdata, styles);
        sga2[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesEmployeeBenefits][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[salesGroupInsurance][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesStatHoliday][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesVacationPay][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesWorkersCompensation][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeEmployeeBenefits][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeGroupInsurance][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officePayrollTaxes][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeVacationPay][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeWorkersCompensation][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeStatHoliday][3].ToString();
        tempp += double.Parse(tempdata, styles);
        sga2[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesEmployeeBenefits][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[salesGroupInsurance][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesStatHoliday][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesVacationPay][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesWorkersCompensation][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeEmployeeBenefits][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeGroupInsurance][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officePayrollTaxes][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeVacationPay][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeWorkersCompensation][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeStatHoliday][5].ToString();
        tempp += double.Parse(tempdata, styles);
        sga2[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesEmployeeBenefits][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[salesGroupInsurance][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesStatHoliday][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesVacationPay][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesWorkersCompensation][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeEmployeeBenefits][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeGroupInsurance][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officePayrollTaxes][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeVacationPay][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeWorkersCompensation][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeStatHoliday][7].ToString();
        tempp += double.Parse(tempdata, styles);
        sga2[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesEmployeeBenefits][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[salesGroupInsurance][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesStatHoliday][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesVacationPay][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[salesWorkersCompensation][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeEmployeeBenefits][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeGroupInsurance][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officePayrollTaxes][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeVacationPay][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeWorkersCompensation][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeStatHoliday][9].ToString();
        tempp += double.Parse(tempdata, styles);
        sga2[9] = NegativeCurrencyFormat(tempp);

        row = table.Rows.Count;

        tempdata = sga2[1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp /= double.Parse(tempdata, styles);
        sga2[2] = tempp.ToString("P2");

        tempdata = sga2[3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp /= double.Parse(tempdata, styles);
        sga2[4] = tempp.ToString("P2");

        tempdata = sga2[5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp /= double.Parse(tempdata, styles);
        sga2[6] = tempp.ToString("P2");

        tempdata = sga2[7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp /= double.Parse(tempdata, styles);
        sga2[8] = tempp.ToString("P2");

        tempdata = sga2[9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][9].ToString();
        tempp /= double.Parse(tempdata, styles);
        sga2[10] = tempp.ToString("P2");

        table.Rows.Add(sga2);

        DataRow sga3 = table.NewRow();
        sga3[0] = "TOTAL S,G&A LABOR COST:";

        row = table.Rows.Count;

        tempdata = table.Rows[row - 2][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp += double.Parse(tempdata, styles);
        sga3[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp += double.Parse(tempdata, styles);
        sga3[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp += double.Parse(tempdata, styles);
        sga3[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp += double.Parse(tempdata, styles);
        sga3[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][9].ToString();
        tempp += double.Parse(tempdata, styles);
        sga3[9] = NegativeCurrencyFormat(tempp);

        tempdata = sga3[1].ToString();
        tempp = double.Parse(tempdata, styles);
        sga3[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = sga3[3].ToString();
        tempp = double.Parse(tempdata, styles);
        sga3[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = sga3[5].ToString();
        tempp = double.Parse(tempdata, styles);
        sga3[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = sga3[7].ToString();
        tempp = double.Parse(tempdata, styles);
        sga3[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = sga3[9].ToString();
        tempp = double.Parse(tempdata, styles);
        sga3[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(sga3);

        blockposition[9, 1] = table.Rows.Count - 1;

        DataRow emp6 = table.NewRow();
        table.Rows.Add(emp6);

        #endregion

        #region Total LABOR
        DataRow labor = table.NewRow();
        labor[0] = "TOTAL WAGES";
        tempdata = table.Rows[blockposition[8, 0]+1][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[9, 0] + 1][1].ToString();
        tempp += double.Parse(tempdata, styles);
        labor[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[8, 0] + 1][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[9, 0] + 1][3].ToString();
        tempp += double.Parse(tempdata, styles);
        labor[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[8, 0] + 1][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[9, 0] + 1][5].ToString();
        tempp += double.Parse(tempdata, styles);
        labor[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[8, 0] + 1][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[9, 0] + 1][7].ToString();
        tempp += double.Parse(tempdata, styles);
        labor[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[8, 0] + 1][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[9, 0] + 1][9].ToString();
        tempp += double.Parse(tempdata, styles);
        labor[9] = NegativeCurrencyFormat(tempp);

        tempdata = labor[1].ToString();
        tempp = double.Parse(tempdata, styles);
        labor[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = labor[3].ToString();
        tempp = double.Parse(tempdata, styles);
        labor[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = labor[5].ToString();
        tempp = double.Parse(tempdata, styles);
        labor[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = labor[7].ToString();
        tempp = double.Parse(tempdata, styles);
        labor[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = labor[9].ToString();
        tempp = double.Parse(tempdata, styles);
        labor[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(labor);

        DataRow labor1 = table.NewRow();
        labor1[0] = "TOTOAL BENEFITS (% IS TO WAGES)";
        tempdata = table.Rows[blockposition[8, 0] + 2][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[9, 0] + 2][1].ToString();
        tempp += double.Parse(tempdata, styles);
        labor1[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[8, 0] + 2][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[9, 0] + 2][3].ToString();
        tempp += double.Parse(tempdata, styles);
        labor1[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[8, 0] + 2][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[9, 0] + 2][5].ToString();
        tempp += double.Parse(tempdata, styles);
        labor1[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[8, 0] + 2][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[9, 0] + 2][7].ToString();
        tempp += double.Parse(tempdata, styles);
        labor1[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[8, 0] + 2][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[9, 0] + 2][9].ToString();
        tempp += double.Parse(tempdata, styles);
        labor1[9] = NegativeCurrencyFormat(tempp);

        row = table.Rows.Count;

        tempdata = labor1[1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp /= double.Parse(tempdata, styles);
        labor1[2] = tempp.ToString("P2");

        tempdata = labor1[3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp /= double.Parse(tempdata, styles);
        labor1[4] = tempp.ToString("P2");

        tempdata = labor1[5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp /= double.Parse(tempdata, styles);
        labor1[6] = tempp.ToString("P2");

        tempdata = labor1[7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp /= double.Parse(tempdata, styles);
        labor1[8] = tempp.ToString("P2");

        tempdata = labor1[9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][9].ToString();
        tempp /= double.Parse(tempdata, styles);
        labor1[10] = tempp.ToString("P2");

        table.Rows.Add(labor1);

        DataRow labor2 = table.NewRow();
        labor2[0] = "TOTAL LABOR COST:";

        row = table.Rows.Count;

        tempdata = table.Rows[row - 2][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp += double.Parse(tempdata, styles);
        labor2[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp += double.Parse(tempdata, styles);
        labor2[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp += double.Parse(tempdata, styles);
        labor2[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp += double.Parse(tempdata, styles);
        labor2[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][9].ToString();
        tempp += double.Parse(tempdata, styles);
        labor2[9] = NegativeCurrencyFormat(tempp);

        tempdata = labor2[1].ToString();
        tempp = double.Parse(tempdata, styles);
        labor2[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = labor2[3].ToString();
        tempp = double.Parse(tempdata, styles);
        labor2[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = labor2[5].ToString();
        tempp = double.Parse(tempdata, styles);
        labor2[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = labor2[7].ToString();
        tempp = double.Parse(tempdata, styles);
        labor2[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = labor2[9].ToString();
        tempp = double.Parse(tempdata, styles);
        labor2[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(labor2);

        totalposition[6] = table.Rows.Count - 1;

        DataRow emp7 = table.NewRow();
        table.Rows.Add(emp7);

        #endregion

        #region Facility costs
        DataRow facost = table.NewRow();
        facost[0] = "FACILITY COSTS";
        table.Rows.Add(facost);

        blockposition[10, 0] = table.Rows.Count - 1;

        DataRow facost1 = table.NewRow();
        facost1[0] = "WAGES";
        tempdata = table.Rows[officeRent][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryRent][1].ToString();
        tempp += double.Parse(tempdata, styles);
        facost1[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeRent][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryRent][3].ToString();
        tempp += double.Parse(tempdata, styles);
        facost1[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeRent][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryRent][5].ToString();
        tempp += double.Parse(tempdata, styles);
        facost1[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeRent][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryRent][7].ToString();
        tempp += double.Parse(tempdata, styles);
        facost1[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeRent][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryRent][9].ToString();
        tempp += double.Parse(tempdata, styles);
        facost1[9] = NegativeCurrencyFormat(tempp);

        tempdata = facost1[1].ToString();
        tempp = double.Parse(tempdata, styles);
        facost1[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = facost1[3].ToString();
        tempp = double.Parse(tempdata, styles);
        facost1[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = facost1[5].ToString();
        tempp = double.Parse(tempdata, styles);
        facost1[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = facost1[7].ToString();
        tempp = double.Parse(tempdata, styles);
        facost1[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = facost1[9].ToString();
        tempp = double.Parse(tempdata, styles);
        facost1[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(facost1);

        DataRow facost2 = table.NewRow();
        facost2[0] = "DEPRECIATION";
        tempdata = table.Rows[factoryDepreBuilding][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryDepreBuildingImprov][1].ToString();
        tempp += double.Parse(tempdata, styles);
        facost2[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreBuilding][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryDepreBuildingImprov][3].ToString();
        tempp += double.Parse(tempdata, styles);
        facost2[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreBuilding][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryDepreBuildingImprov][5].ToString();
        tempp += double.Parse(tempdata, styles);
        facost2[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreBuilding][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryDepreBuildingImprov][7].ToString();
        tempp += double.Parse(tempdata, styles);
        facost2[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreBuilding][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryDepreBuildingImprov][9].ToString();
        tempp += double.Parse(tempdata, styles);
        facost2[9] = NegativeCurrencyFormat(tempp);

        tempdata = facost2[1].ToString();
        tempp = double.Parse(tempdata, styles);
        facost2[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = facost2[3].ToString();
        tempp = double.Parse(tempdata, styles);
        facost2[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = facost2[5].ToString();
        tempp = double.Parse(tempdata, styles);
        facost2[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = facost2[7].ToString();
        tempp = double.Parse(tempdata, styles);
        facost2[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = facost2[9].ToString();
        tempp = double.Parse(tempdata, styles);
        facost2[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(facost2);

        DataRow facost3 = table.NewRow();
        facost3[0] = "REPAIR & MAINTENANCE";
        tempdata = table.Rows[factoryReqMainBuilding][1].ToString();
        tempp = double.Parse(tempdata, styles);
        facost3[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryReqMainBuilding][3].ToString();
        tempp = double.Parse(tempdata, styles);
        facost3[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryReqMainBuilding][5].ToString();
        tempp = double.Parse(tempdata, styles);
        facost3[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryReqMainBuilding][7].ToString();
        tempp = double.Parse(tempdata, styles);
        facost3[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryReqMainBuilding][9].ToString();
        tempp = double.Parse(tempdata, styles);
        facost3[9] = NegativeCurrencyFormat(tempp);

        tempdata = facost3[1].ToString();
        tempp = double.Parse(tempdata, styles);
        facost3[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = facost3[3].ToString();
        tempp = double.Parse(tempdata, styles);
        facost3[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = facost3[5].ToString();
        tempp = double.Parse(tempdata, styles);
        facost3[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = facost3[7].ToString();
        tempp = double.Parse(tempdata, styles);
        facost3[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = facost3[9].ToString();
        tempp = double.Parse(tempdata, styles);
        facost3[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(facost3);

        DataRow facost4 = table.NewRow();
        facost4[0] = "PROPERTY TAX";
        tempdata = table.Rows[factoryPropertyTax][1].ToString();
        tempp = double.Parse(tempdata, styles);
        facost4[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryPropertyTax][3].ToString();
        tempp = double.Parse(tempdata, styles);
        facost4[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryPropertyTax][5].ToString();
        tempp = double.Parse(tempdata, styles);
        facost4[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryPropertyTax][7].ToString();
        tempp = double.Parse(tempdata, styles);
        facost4[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryPropertyTax][9].ToString();
        tempp = double.Parse(tempdata, styles);
        facost4[9] = NegativeCurrencyFormat(tempp);

        tempdata = facost4[1].ToString();
        tempp = double.Parse(tempdata, styles);
        facost4[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = facost4[3].ToString();
        tempp = double.Parse(tempdata, styles);
        facost4[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = facost4[5].ToString();
        tempp = double.Parse(tempdata, styles);
        facost4[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = facost4[7].ToString();
        tempp = double.Parse(tempdata, styles);
        facost4[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = facost4[9].ToString();
        tempp = double.Parse(tempdata, styles);
        facost4[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(facost4);

        row = table.Rows.Count;

        DataRow facost5 = table.NewRow();
        facost5[0] = "TOTAL FACILITY COST:";
        tempdata = table.Rows[row - 4][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp += double.Parse(tempdata, styles);
        facost5[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 4][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp += double.Parse(tempdata, styles);
        facost5[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 4][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp += double.Parse(tempdata, styles);
        facost5[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 4][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp += double.Parse(tempdata, styles);
        facost5[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 4][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][9].ToString();
        tempp += double.Parse(tempdata, styles);
        facost5[9] = NegativeCurrencyFormat(tempp);

        tempdata = facost5[1].ToString();
        tempp = double.Parse(tempdata, styles);
        facost5[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = facost5[3].ToString();
        tempp = double.Parse(tempdata, styles);
        facost5[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = facost5[5].ToString();
        tempp = double.Parse(tempdata, styles);
        facost5[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = facost5[7].ToString();
        tempp = double.Parse(tempdata, styles);
        facost5[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = facost5[9].ToString();
        tempp = double.Parse(tempdata, styles);
        facost5[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(facost5);

        totalposition[7] = table.Rows.Count - 1;

        DataRow emp8 = table.NewRow();
        table.Rows.Add(emp8);
        #endregion

        #region DEPRECIATION
        DataRow depreciation = table.NewRow();
        depreciation[0] = "DEPRECIATION";
        table.Rows.Add(depreciation);

        blockposition[11, 0] = table.Rows.Count - 1;

        DataRow depreciation1 = table.NewRow();
        depreciation1[0] = "BUILDING";
        tempdata = table.Rows[factoryDepreBuilding][1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation1[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreBuilding][3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation1[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreBuilding][5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation1[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreBuilding][7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation1[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreBuilding][9].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation1[9] = NegativeCurrencyFormat(tempp);

        tempdata = depreciation1[1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation1[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = depreciation1[3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation1[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = depreciation1[5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation1[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = depreciation1[7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation1[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = depreciation1[9].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation1[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(depreciation1);

        DataRow depreciation2 = table.NewRow();
        depreciation2[0] = "BUILDING IMPROVEMENT";
        tempdata = table.Rows[factoryDepreBuildingImprov][1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation2[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreBuildingImprov][3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation2[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreBuildingImprov][5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation2[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreBuildingImprov][7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation2[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreBuildingImprov][9].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation2[9] = NegativeCurrencyFormat(tempp);

        tempdata = depreciation2[1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation2[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = depreciation2[3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation2[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = depreciation2[5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation2[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = depreciation2[7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation2[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = depreciation2[9].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation2[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(depreciation2);

        row = table.Rows.Count;

        DataRow depreciation3 = table.NewRow();
        depreciation3[0] = "TOTAL BUILDING:";
        tempdata = table.Rows[row - 2][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation3[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation3[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation3[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation3[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][9].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation3[9] = NegativeCurrencyFormat(tempp);

        tempdata = depreciation3[1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation3[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = depreciation3[3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation3[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = depreciation3[5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation3[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = depreciation3[7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation3[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = depreciation3[9].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation3[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(depreciation3);
        blockposition[11, 1] = table.Rows.Count - 1;

        DataRow emp9 = table.NewRow();
        table.Rows.Add(emp9);

        blockposition[12, 0] = table.Rows.Count - 1;

        DataRow depreciation4 = table.NewRow();
        depreciation4[0] = "MACHINERY & EQUIPMENT";
        tempdata = table.Rows[factoryDepreMachine][1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation4[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreMachine][3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation4[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreMachine][5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation4[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreMachine][7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation4[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreMachine][9].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation4[9] = NegativeCurrencyFormat(tempp);

        tempdata = depreciation4[1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation4[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = depreciation4[3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation4[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = depreciation4[5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation4[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = depreciation4[7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation4[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = depreciation4[9].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation4[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(depreciation4);

        DataRow depreciation5 = table.NewRow();
        depreciation5[0] = "FURNITURE & FIXTURES, COMPUTERS/SOFTWARE";
        tempdata = table.Rows[factoryDepreFurniture][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryDepreSoftware][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeDepreFurniture][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeSoftwareAmort][1].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation5[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreFurniture][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryDepreSoftware][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeDepreFurniture][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeSoftwareAmort][3].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation5[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreFurniture][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryDepreSoftware][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeDepreFurniture][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeSoftwareAmort][5].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation5[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryDepreFurniture][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[factoryDepreSoftware][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeDepreFurniture][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[officeSoftwareAmort][7].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation5[7] = NegativeCurrencyFormat(tempp);

        //tempdata = table.Rows[factoryDepreFurniture][9].ToString();
        //tempp = double.Parse(tempdata, styles);
        //tempdata = table.Rows[factoryDepreSoftware][9].ToString();
        //tempp += double.Parse(tempdata, styles);
        //tempdata = table.Rows[officeDepreFurniture][9].ToString();
        //tempp += double.Parse(tempdata, styles);
        //tempdata = table.Rows[officeSoftwareAmort][9].ToString();
        //tempp += double.Parse(tempdata, styles);
        //depreciation5[9] = NegativeCurrencyFormat(tempp);

        tempdata = depreciation5[1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation5[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = depreciation5[3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation5[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = depreciation5[5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation5[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = depreciation5[7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation5[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = depreciation5[9].ToString();
        //tempp = double.Parse(tempdata, styles);
        //depreciation5[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(depreciation5);

        DataRow depreciation6 = table.NewRow();
        depreciation6[0] = "VEHICLES";
        tempdata = table.Rows[salesDepreAuto][1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation6[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesDepreAuto][3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation6[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesDepreAuto][5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation6[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesDepreAuto][7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation6[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesDepreAuto][9].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation6[9] = NegativeCurrencyFormat(tempp);

        tempdata = depreciation6[1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation6[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = depreciation6[3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation6[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = depreciation6[5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation6[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = depreciation6[7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation6[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = depreciation6[9].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation6[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(depreciation6);

        DataRow depreciation8 = table.NewRow();
        depreciation8[0] = "TOOLING";
        tempdata = table.Rows[factoryToolAmort][1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation8[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryToolAmort][3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation8[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryToolAmort][5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation8[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryToolAmort][7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation8[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryToolAmort][9].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation8[9] = NegativeCurrencyFormat(tempp);

        tempdata = depreciation8[1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation8[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = depreciation8[3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation8[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = depreciation8[5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation8[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = depreciation8[7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation8[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = depreciation8[9].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation8[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(depreciation8);

        row = table.Rows.Count;

        DataRow depreciation9 = table.NewRow();
        depreciation9[0] = "TOTAL NON-BUILDING DEPRECIATION:";
        tempdata = table.Rows[row - 4][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation9[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 4][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation9[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 4][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation9[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 4][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation9[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 4][9].ToString();
        tempp = double.Parse(tempdata, styles);
        //tempdata = table.Rows[row - 3][9].ToString(); depreciation values
        //tempp += double.Parse(tempdata, styles);
        //tempdata = table.Rows[row - 2][9].ToString();
        //tempp += double.Parse(tempdata, styles);
        //tempdata = table.Rows[row - 1][9].ToString();
        //tempp += double.Parse(tempdata, styles);
        depreciation9[9] = NegativeCurrencyFormat(tempp);

        tempdata = depreciation9[1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation9[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = depreciation9[3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation9[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = depreciation9[5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation9[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = depreciation9[7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation9[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = depreciation9[9].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation9[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(depreciation9);

        blockposition[12, 1] = table.Rows.Count - 1;
        DataRow emp10 = table.NewRow();
        table.Rows.Add(emp10);

        DataRow depreciation11 = table.NewRow();
        depreciation11[0] = "TOTAL DEPRECIATION:";
        tempdata = table.Rows[blockposition[11, 1]][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[12, 1]][1].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation11[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[11, 1]][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[12, 1]][3].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation11[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[11, 1]][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[12, 1]][5].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation11[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[11, 1]][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[12, 1]][7].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation11[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[11, 1]][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[12, 1]][9].ToString();
        tempp += double.Parse(tempdata, styles);
        depreciation11[9] = NegativeCurrencyFormat(tempp);

        tempdata = depreciation11[1].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation11[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = depreciation11[3].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation11[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = depreciation11[5].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation11[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = depreciation11[7].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation11[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = depreciation11[9].ToString();
        tempp = double.Parse(tempdata, styles);
        depreciation11[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(depreciation11);

        totalposition[8] = table.Rows.Count - 1;

        DataRow emp11 = table.NewRow();
        table.Rows.Add(emp11);

        #endregion

        #region TOOLING EXPENSE
        DataRow tool = table.NewRow();
        tool[0] = "TOOLING EXPENSE";
        table.Rows.Add(tool);

        blockposition[13, 0] = table.Rows.Count - 1;

        DataRow tool1 = table.NewRow();
        tool1[0] = "TOOLING EXPENSE";
        tempdata = table.Rows[factoryToolExpense][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tool1[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryToolExpense][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tool1[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryToolExpense][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tool1[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryToolExpense][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tool1[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryToolExpense][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tool1[9] = NegativeCurrencyFormat(tempp);

        tempdata = tool1[1].ToString();
        tempp = double.Parse(tempdata, styles);
        tool1[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = tool1[3].ToString();
        tempp = double.Parse(tempdata, styles);
        tool1[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = tool1[5].ToString();
        tempp = double.Parse(tempdata, styles);
        tool1[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = tool1[7].ToString();
        tempp = double.Parse(tempdata, styles);
        tool1[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = tool1[9].ToString();
        tempp = double.Parse(tempdata, styles);
        tool1[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(tool1);

        DataRow tool2 = table.NewRow();
        tool2[0] = "TOOLING AMORTIZATION";
        tempdata = table.Rows[factoryToolAmort][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tool2[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryToolAmort][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tool2[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryToolAmort][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tool2[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryToolAmort][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tool2[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryToolAmort][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tool2[9] = NegativeCurrencyFormat(tempp);

        tempdata = tool2[1].ToString();
        tempp = double.Parse(tempdata, styles);
        tool2[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = tool2[3].ToString();
        tempp = double.Parse(tempdata, styles);
        tool2[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = tool2[5].ToString();
        tempp = double.Parse(tempdata, styles);
        tool2[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = tool2[7].ToString();
        tempp = double.Parse(tempdata, styles);
        tool2[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = tool2[9].ToString();
        tempp = double.Parse(tempdata, styles);
        tool2[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(tool2);
        blockposition[13, 1] = table.Rows.Count - 1;

        row = table.Rows.Count;
        DataRow tool3 = table.NewRow();
        tool3[0] = "TOTAL TOOLING EXPENSE:";
        tempdata = table.Rows[row - 2][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tool3[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tool3[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tool3[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tool3[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 2][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tool3[9] = NegativeCurrencyFormat(tempp);

        tempdata = tool3[1].ToString();
        tempp = double.Parse(tempdata, styles);
        tool3[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = tool3[3].ToString();
        tempp = double.Parse(tempdata, styles);
        tool3[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = tool3[5].ToString();
        tempp = double.Parse(tempdata, styles);
        tool3[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = tool3[7].ToString();
        tempp = double.Parse(tempdata, styles);
        tool3[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = tool3[9].ToString();
        tempp = double.Parse(tempdata, styles);
        tool3[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(tool3);

        totalposition[9] = table.Rows.Count - 1;

        DataRow emp12 = table.NewRow();
        table.Rows.Add(emp12);

        #endregion

        #region supplies
        DataRow supp = table.NewRow();
        supp[0] = "SUPPLIES";
        table.Rows.Add(supp);

        blockposition[14, 0] = table.Rows.Count - 1;

        DataRow supp1 = table.NewRow();
        supp1[0] = "CAD/CAM SUPPLIES";
        tempdata = table.Rows[factoryCADCAMSupplies][1].ToString();
        tempp = double.Parse(tempdata, styles);
        supp1[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryCADCAMSupplies][3].ToString();
        tempp = double.Parse(tempdata, styles);
        supp1[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryCADCAMSupplies][5].ToString();
        tempp = double.Parse(tempdata, styles);
        supp1[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryCADCAMSupplies][7].ToString();
        tempp = double.Parse(tempdata, styles);
        supp1[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryCADCAMSupplies][9].ToString();
        tempp = double.Parse(tempdata, styles);
        supp1[9] = NegativeCurrencyFormat(tempp);

        tempdata = supp1[1].ToString();
        tempp = double.Parse(tempdata, styles);
        supp1[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = supp1[3].ToString();
        tempp = double.Parse(tempdata, styles);
        supp1[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = supp1[5].ToString();
        tempp = double.Parse(tempdata, styles);
        supp1[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = supp1[7].ToString();
        tempp = double.Parse(tempdata, styles);
        supp1[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = supp1[9].ToString();
        tempp = double.Parse(tempdata, styles);
        supp1[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(supp1);

        DataRow supp2 = table.NewRow();
        supp2[0] = "SHOP SUPPLIES";
        tempdata = table.Rows[factoryShopSupplies][1].ToString();
        tempp = double.Parse(tempdata, styles);
        supp2[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryShopSupplies][3].ToString();
        tempp = double.Parse(tempdata, styles);
        supp2[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryShopSupplies][5].ToString();
        tempp = double.Parse(tempdata, styles);
        supp2[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryShopSupplies][7].ToString();
        tempp = double.Parse(tempdata, styles);
        supp2[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryShopSupplies][9].ToString();
        tempp = double.Parse(tempdata, styles);
        supp2[9] = NegativeCurrencyFormat(tempp);

        tempdata = supp2[1].ToString();
        tempp = double.Parse(tempdata, styles);
        supp2[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = supp2[3].ToString();
        tempp = double.Parse(tempdata, styles);
        supp2[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = supp2[5].ToString();
        tempp = double.Parse(tempdata, styles);
        supp2[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = supp2[7].ToString();
        tempp = double.Parse(tempdata, styles);
        supp2[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = supp2[9].ToString();
        tempp = double.Parse(tempdata, styles);
        supp2[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(supp2);

        DataRow supp3 = table.NewRow();
        supp3[0] = "HEAT TREAT SUPPLIES";
        tempdata = table.Rows[factoryHeatTreatSupplies][1].ToString();
        tempp = double.Parse(tempdata, styles);
        supp3[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryHeatTreatSupplies][3].ToString();
        tempp = double.Parse(tempdata, styles);
        supp3[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryHeatTreatSupplies][5].ToString();
        tempp = double.Parse(tempdata, styles);
        supp3[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryHeatTreatSupplies][7].ToString();
        tempp = double.Parse(tempdata, styles);
        supp3[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryHeatTreatSupplies][9].ToString();
        tempp = double.Parse(tempdata, styles);
        supp3[9] = NegativeCurrencyFormat(tempp);

        tempdata = supp3[1].ToString();
        tempp = double.Parse(tempdata, styles);
        supp3[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = supp3[3].ToString();
        tempp = double.Parse(tempdata, styles);
        supp3[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = supp3[5].ToString();
        tempp = double.Parse(tempdata, styles);
        supp3[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = supp3[7].ToString();
        tempp = double.Parse(tempdata, styles);
        supp3[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = supp3[9].ToString();
        tempp = double.Parse(tempdata, styles);
        supp3[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(supp3);

        DataRow supp4 = table.NewRow();
        supp4[0] = "OFFICE SUPPLIES";
        tempdata = table.Rows[officeSupplies][1].ToString();
        tempp = double.Parse(tempdata, styles);
        supp4[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeSupplies][3].ToString();
        tempp = double.Parse(tempdata, styles);
        supp4[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeSupplies][5].ToString();
        tempp = double.Parse(tempdata, styles);
        supp4[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeSupplies][7].ToString();
        tempp = double.Parse(tempdata, styles);
        supp4[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeSupplies][9].ToString();
        tempp = double.Parse(tempdata, styles);
        supp4[9] = NegativeCurrencyFormat(tempp);

        tempdata = supp4[1].ToString();
        tempp = double.Parse(tempdata, styles);
        supp4[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = supp4[3].ToString();
        tempp = double.Parse(tempdata, styles);
        supp4[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = supp4[5].ToString();
        tempp = double.Parse(tempdata, styles);
        supp4[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = supp4[7].ToString();
        tempp = double.Parse(tempdata, styles);
        supp4[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = supp4[9].ToString();
        tempp = double.Parse(tempdata, styles);
        supp4[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(supp4);
        blockposition[14, 1] = table.Rows.Count;

        row = table.Rows.Count;
        DataRow supp5 = table.NewRow();
        supp5[0] = "TOTAL SUPPLIES:";
        tempdata = table.Rows[row - 4][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp += double.Parse(tempdata, styles);
        supp5[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 4][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp += double.Parse(tempdata, styles);
        supp5[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 4][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp += double.Parse(tempdata, styles);
        supp5[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 4][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp += double.Parse(tempdata, styles);
        supp5[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 4][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][9].ToString();
        tempp += double.Parse(tempdata, styles); 
        tempdata = table.Rows[row - 2][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][9].ToString();
        tempp += double.Parse(tempdata, styles);
        supp5[9] = NegativeCurrencyFormat(tempp);

        tempdata = supp5[1].ToString();
        tempp = double.Parse(tempdata, styles);
        supp5[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = supp5[3].ToString();
        tempp = double.Parse(tempdata, styles);
        supp5[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = supp5[5].ToString();
        tempp = double.Parse(tempdata, styles);
        supp5[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = supp5[7].ToString();
        tempp = double.Parse(tempdata, styles);
        supp5[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = supp5[9].ToString();
        tempp = double.Parse(tempdata, styles);
        supp5[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(supp5);

        totalposition[10] = table.Rows.Count - 1;

        DataRow emp13 = table.NewRow();
        table.Rows.Add(emp13);
        #endregion

        #region TRAVEL, MEALS & ENTERTAINMENT
        DataRow travel = table.NewRow();
        travel[0] = "TRAVEL, MEALS & ENTERTAINMENT";
        table.Rows.Add(travel);

        blockposition[15, 0] = table.Rows.Count - 1;

        DataRow travel1 = table.NewRow();
        travel1[0] = "TRAVEL - SHOP";
        tempdata = table.Rows[factoryTravel][1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel1[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryTravel][3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel1[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryTravel][5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel1[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryTravel][7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel1[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryTravel][9].ToString();
        //tempp = double.Parse(tempdata, styles);
        travel1[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel1[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel1[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel1[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel1[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel1[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel1[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel1[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel1[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel1[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel1[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel1);

        DataRow travel2 = table.NewRow();
        travel2[0] = "MEALS & ENTERTAINMENT - SHOP";
        tempdata = table.Rows[factoryMeals][1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel2[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryMeals][3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel2[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryMeals][5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel2[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryMeals][7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel2[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryMeals][9].ToString();
        //tempp = double.Parse(tempdata, styles);
        travel2[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel2[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel2[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel2[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel2[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel2[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel2[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel2[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel2[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel2[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel2[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel2);

        DataRow travel3 = table.NewRow();
        travel3[0] = "AIRFARE - SHOP";
        tempdata = table.Rows[factoryAirFare][1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel3[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryAirFare][3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel3[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryAirFare][5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel3[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryAirFare][7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel3[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factoryAirFare][9].ToString();
       // tempp = double.Parse(tempdata, styles);
        travel3[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel3[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel3[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel3[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel3[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel3[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel3[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel3[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel3[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel3[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel3[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel3);

        DataRow travel4 = table.NewRow();
        travel4[0] = "SOUTH AMERICA TRAVEL - SHOP";
        tempdata = table.Rows[factorySATravelRow][1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel4[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factorySATravelRow][3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel4[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factorySATravelRow][5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel4[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factorySATravelRow][7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel4[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[factorySATravelRow][9].ToString();
        //tempp = double.Parse(tempdata, styles);
        travel4[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel4[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel4[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel4[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel4[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel4[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel4[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel4[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel4[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel4[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel4[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel4);

        DataRow travel5 = table.NewRow();
        travel5[0] = "NON DEDUCTIBLE EXP GOLF";
        tempdata = table.Rows[salesGolf][1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel5[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesGolf][3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel5[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesGolf][5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel5[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesGolf][7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel5[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesGolf][9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel5[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel5[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel5[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel5[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel5[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel5[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel5[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel5[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel5[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel5[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel5[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel5);

        DataRow travel6 = table.NewRow();
        travel6[0] = "SELLING & TRAVEL MEALS & ENTERTAINMENT";
        tempdata = table.Rows[salesMeals][1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel6[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesMeals][3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel6[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesMeals][5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel6[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesMeals][7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel6[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesMeals][9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel6[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel6[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel6[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel6[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel6[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel6[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel6[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel6[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel6[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel6[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel6[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel6);

        DataRow travel7 = table.NewRow();
        travel7[0] = "SELLING & TRAVEL EXPENSES";
        tempdata = table.Rows[salesTravel][1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel7[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesTravel][3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel7[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesTravel][5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel7[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesTravel][7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel7[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesTravel][9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel7[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel7[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel7[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel7[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel7[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel7[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel7[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel7[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel7[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel7[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel7[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel7);

        DataRow travel8 = table.NewRow();
        travel8[0] = "SELLING & TRAVEL AIR FARE";
        tempdata = table.Rows[salesAirFare][1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel8[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesAirFare][3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel8[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesAirFare][5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel8[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesAirFare][7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel8[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[salesAirFare][9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel8[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel8[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel8[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel8[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel8[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel8[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel8[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel8[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel8[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel8[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel8[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel8);

        DataRow travel9 = table.NewRow();
        travel9[0] = "OFFICE MEALS & ENTERTAINMENT";
        tempdata = table.Rows[officeMeals][1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel9[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeMeals][3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel9[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeMeals][5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel9[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeMeals][7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel9[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeMeals][9].ToString();
        //tempp = double.Parse(tempdata, styles);
        travel9[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel9[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel9[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel9[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel9[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel9[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel9[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel9[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel9[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel9[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel9[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel9);

        DataRow travel10 = table.NewRow();
        travel10[0] = "OFFICE TRAVEL EXPENSES";
        tempdata = table.Rows[officeTravel][1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel10[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeTravel][3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel10[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeTravel][5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel10[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeTravel][7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel10[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeTravel][9].ToString();
        //tempp = double.Parse(tempdata, styles);
        travel10[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel10[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel10[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel10[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel10[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel10[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel10[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel10[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel10[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel10[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel10[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel10);

        DataRow travel11 = table.NewRow();
        travel11[0] = "OFFICE AIR FARE";
        tempdata = table.Rows[officeAirFare][1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel11[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeAirFare][3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel11[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeAirFare][5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel11[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeAirFare][7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel11[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeAirFare][9].ToString();
        //tempp = double.Parse(tempdata, styles);
        travel11[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel11[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel11[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel11[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel11[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel11[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel11[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel11[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel11[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel11[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel11[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel11);

        DataRow travel12 = table.NewRow();
        travel12[0] = "OFFICE SOUTH AMERICA TRAVEL";
        tempdata = table.Rows[officeSATravelRow][1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel12[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeSATravelRow][3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel12[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeSATravelRow][5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel12[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeSATravelRow][7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel12[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[officeSATravelRow][9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel12[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel12[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel12[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel12[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel12[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel12[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel12[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel12[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel12[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel12[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel12[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel12);

        DataRow travel13 = table.NewRow();
        travel13[0] = "RECLASS SOUTH AMERICA TRAVEL";
        tempdata = table.Rows[blockposition[15, 0] + 4][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[15, 0] + 12][1].ToString();
        tempp -= double.Parse(tempdata, styles);
        travel13[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[15, 0] + 4][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[15, 0] + 12][3].ToString();
        tempp -= double.Parse(tempdata, styles);
        travel13[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[15, 0] + 4][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[15, 0] + 12][5].ToString();
        tempp -= double.Parse(tempdata, styles);
        travel13[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[15, 0] + 4][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[15, 0] + 12][7].ToString();
        tempp -= double.Parse(tempdata, styles);
        travel13[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[blockposition[15, 0] + 4][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[blockposition[15, 0] + 12][9].ToString();
        tempp -= double.Parse(tempdata, styles);
        travel13[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel13[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel13[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel13[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel13[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel13[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel13[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel13[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel13[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel13[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel13[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel13);

        blockposition[15, 1] = table.Rows.Count;

        row = table.Rows.Count;
        DataRow travel14 = table.NewRow();
        travel14[0] = "TOTAL TRAVEL, MEALS & ENTERTAINMENT:";
        tempdata = table.Rows[row - 13][1].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 12][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 11][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 10][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 9][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 8][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 7][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 6][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 5][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 4][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][1].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][1].ToString();
        tempp += double.Parse(tempdata, styles);
        travel14[1] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 13][3].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 12][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 11][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 10][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 9][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 8][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 7][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 6][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 5][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 4][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][3].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][3].ToString();
        tempp += double.Parse(tempdata, styles);
        travel14[3] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 13][5].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 12][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 11][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 10][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 9][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 8][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 7][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 6][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 5][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 4][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][5].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][5].ToString();
        tempp += double.Parse(tempdata, styles);
        travel14[5] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 13][7].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 12][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 11][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 10][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 9][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 8][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 7][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 6][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 5][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 4][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][7].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][7].ToString();
        tempp += double.Parse(tempdata, styles);
        travel14[7] = NegativeCurrencyFormat(tempp);

        tempdata = table.Rows[row - 13][9].ToString();
        tempp = double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 12][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 11][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 10][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 9][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 8][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 7][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 6][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 5][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 4][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 3][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 2][9].ToString();
        tempp += double.Parse(tempdata, styles);
        tempdata = table.Rows[row - 1][9].ToString();
        tempp += double.Parse(tempdata, styles);
        travel14[9] = NegativeCurrencyFormat(tempp);

        tempdata = travel14[1].ToString();
        tempp = double.Parse(tempdata, styles);
        travel14[2] = (tempp / allsummary[0, 0]).ToString("P2");

        tempdata = travel14[3].ToString();
        tempp = double.Parse(tempdata, styles);
        travel14[4] = (tempp / allsummary[0, 2]).ToString("P2");

        tempdata = travel14[5].ToString();
        tempp = double.Parse(tempdata, styles);
        travel14[6] = (tempp / allsummary[0, 4]).ToString("P2");

        tempdata = travel14[7].ToString();
        tempp = double.Parse(tempdata, styles);
        travel14[8] = (tempp / allsummary[0, 6]).ToString("P2");

        tempdata = travel14[9].ToString();
        tempp = double.Parse(tempdata, styles);
        travel14[10] = (tempp / allsummary[0, 8]).ToString("P2");

        table.Rows.Add(travel14);

        totalposition[11] = table.Rows.Count - 1;

        DataRow emp14 = table.NewRow();
        table.Rows.Add(emp14);
        #endregion

        table.AcceptChanges();
        gridView.DataSource = table;
        gridView.DataBind();

        TableStyleAdjustmentConsolidate(table, gridView);
    }

    protected void MichiganReport()
    {
        Session["_plantID"] = "003";
        Session["_plantName"] = "Michigan";
        currency = "US";
        GridView gridView = GridViewIncomeStatementReport;
        ExcoCalendar calendar = new ExcoCalendar(fiscalYear, fiscalMonth, true, 1);
        DateTime datename = new DateTime(fiscalYear + 2000, calendar.GetCalendarMonth(), 1);
        string monthname = datename.ToString("MMM");
        string yearname;
        if (monthname == "Oct" || monthname == "Nov" || monthname == "Dec")
        {
            yearname = (fiscalYear + 1999).ToString();
        }
        else
        {
            yearname = (fiscalYear + 2000).ToString();
        }
        gridView.Caption = "<font size=\"9\" color=\"red\">Michigan Income Statement Report (USD) Period " + fiscalMonth + " - " + (fiscalYear + 2000) + " Ending " + monthname + " " + DateTime.DaysInMonth(fiscalYear + 2000, calendar.GetCalendarMonth()) + ", " + yearname + "</font>";
        PanelIncomeStatementReport.Visible = true;
        gridView.Visible = true;
        //write table data
        DataTable table = new DataTable();
        //header column
        WriteHeaderByPlant(table);

        //call .dll to process all the table data
        Process process = new Process(fiscalYear, fiscalMonth);

        //Block 1 : Sales                     
        WriteAmount("SALES", table, 3, process.ss.groupList, 0, currency);

        //Block 2 : Cost of Steels
        WriteAmount("COST OF STEELS", table, 3, process.cs.groupList, 1, currency);

        //Block 3 : Cost of Steels
        WriteAmount("DIRECT LABOUR", table, 3, process.dl.groupList, 2, currency);

        //Block 4 : Factory Overhead
        WriteAmount("FACTORY OVERHEAD", table, 3, process.fo.groupList, 3, currency);

        //Total cost of goods
        WriteTotalLine(table, 0);

        //Gross Margin
        WriteTotalLine(table, 1);

        //Block 5 : Delivery and selling
        WriteAmount("DELIVERY AND SELLING", table, 3, process.ds.groupList, 4, currency);

        //Block 6 : General and administration
        WriteAmount("GENERAL AND ADMINISTRATION", table, 3, process.ga.groupList, 5, currency);

        //Block 7 : Other expenses and income
        WriteAmount("OTHER EXPENSES AND INCOME", table, 3, process.oe.groupList, 6, currency);

        //Total expenses
        WriteTotalLine(table, 2);

        //net income before provision tax
        WriteTotalLine(table, 3);

        //Block 8 : Provisional Taxes
        WriteAmount("", table, 3, process.pt.groupList, 7, currency);

        //Total net income
        WriteTotalLine(table, 4);

        table.AcceptChanges();
        gridView.DataSource = table;
        gridView.DataBind();

        //style adjust
        TableStyleAdjustment(table, gridView);
    }

    protected void ColombiaReport()
    {
        Session["_plantID"] = "004";
        Session["_plantName"] = "Colombia";
        currency = "CP";
        GridView gridView = GridViewIncomeStatementReport;
        ExcoCalendar calendar = new ExcoCalendar(fiscalYear, fiscalMonth, true, 1);
        DateTime datename = new DateTime(fiscalYear + 2000, calendar.GetCalendarMonth(), 1);
        string monthname = datename.ToString("MMM");
        string yearname;
        if (monthname == "Oct" || monthname == "Nov" || monthname == "Dec")
        {
            yearname = (fiscalYear + 1999).ToString();
        }
        else
        {
            yearname = (fiscalYear + 2000).ToString();
        }

        gridView.Caption = "<font size=\"9\" color=\"red\">Colombia Income Statement Report (PESO): Period " + fiscalMonth + " - " + (fiscalYear + 2000) + " Ending " + monthname + " " + DateTime.DaysInMonth(fiscalYear + 2000, calendar.GetCalendarMonth()) + ", " + yearname + "</font>";
        PanelIncomeStatementReport.Visible = true;
        gridView.Visible = true;
        //write table data
        DataTable table = new DataTable();
        //header column
        WriteHeaderByPlant(table);

        //call .dll to process all the table data
        Process process = new Process(fiscalYear, fiscalMonth);

        //Block 1 : Sales                     
        WriteAmount("SALES", table, 4, process.ss.groupList, 0, currency);

        //Block 2 : Cost of Steels
        WriteAmount("COST OF STEELS", table, 4, process.cs.groupList, 1, currency);

        //Block 3 : Cost of Steels
        WriteAmount("DIRECT LABOUR", table, 4, process.dl.groupList, 2, currency);

        //Block 4 : Factory Overhead
        WriteAmount("FACTORY OVERHEAD", table, 4, process.fo.groupList, 3, currency);

        //Total cost of goods
        WriteTotalLine(table, 0);

        //Gross Margin
        WriteTotalLine(table, 1);

        //Block 5 : Delivery and selling
        WriteAmount("DELIVERY AND SELLING", table, 4, process.ds.groupList, 4, currency);

        //Block 6 : General and administration
        WriteAmount("GENERAL AND ADMINISTRATION", table, 4, process.ga.groupList, 5, currency);

        //Block 7 : Other expenses and income
        WriteAmount("OTHER EXPENSES AND INCOME", table, 4, process.oe.groupList, 6, currency);

        //Total expenses
        WriteTotalLine(table, 2);

        //net income before provision tax
        WriteTotalLine(table, 3);

        //Block 8 : Provisional Taxes
        WriteAmount("", table, 4, process.pt.groupList, 7, currency);

        //Total net income
        WriteTotalLine(table, 4);

        table.AcceptChanges();
        gridView.DataSource = table;
        gridView.DataBind();

        //style adjust
        TableStyleAdjustment(table, gridView); 
    }

    protected void TexasReport()
    {
        Session["_plantID"] = "005";
        Session["_plantName"] = "Texas";
        currency = "US";
        GridView gridView = GridViewIncomeStatementReport;
        ExcoCalendar calendar = new ExcoCalendar(fiscalYear, fiscalMonth, true, 1);
        DateTime datename = new DateTime(fiscalYear + 2000, calendar.GetCalendarMonth(), 1);
        string monthname = datename.ToString("MMM");
        string yearname;
        if (monthname == "Oct" || monthname == "Nov" || monthname == "Dec")
        {
            yearname = (fiscalYear + 1999).ToString();
        }
        else
        {
            yearname = (fiscalYear + 2000).ToString();
        }
        gridView.Caption = "<font size=\"9\" color=\"red\">Texas Income Statement Report (USD): Period " + fiscalMonth + " - " + (fiscalYear + 2000) + " Ending " + monthname + " " + DateTime.DaysInMonth(fiscalYear + 2000, calendar.GetCalendarMonth()) + ", " + yearname + "</font>";
        PanelIncomeStatementReport.Visible = true;
        gridView.Visible = true;
        //write table data
        DataTable table = new DataTable();        
        //header column
        WriteHeaderByPlant(table);

        //call .dll to process all the table data
        Process process = new Process(fiscalYear, fiscalMonth);

        //Block 1 : Sales                     
        WriteAmount("SALES", table, 5, process.ss.groupList, 0, currency);

        //Block 2 : Cost of Steels
        WriteAmount("COST OF STEELS", table, 5, process.cs.groupList, 1, currency);

        //Block 3 : Cost of Steels
        WriteAmount("DIRECT LABOUR", table, 5, process.dl.groupList, 2, currency);

        //Block 4 : Factory Overhead
        WriteAmount("FACTORY OVERHEAD", table, 5, process.fo.groupList, 3, currency);

        //Total cost of goods
        WriteTotalLine(table, 0);

        //Gross Margin
        WriteTotalLine(table, 1);

        //Block 5 : Delivery and selling
        WriteAmount("DELIVERY AND SELLING", table, 5, process.ds.groupList, 4, currency);

        //Block 6 : General and administration
        WriteAmount("GENERAL AND ADMINISTRATION", table, 5, process.ga.groupList, 5, currency);

        //Block 7 : Other expenses and income
        WriteAmount("OTHER EXPENSES AND INCOME", table, 5, process.oe.groupList, 6, currency);

        //Total expenses
        WriteTotalLine(table, 2);

        //net income before provision tax
        WriteTotalLine(table, 3);

        //Block 8 : Provisional Taxes
        WriteAmount("", table, 5, process.pt.groupList, 7, currency);

        //Total net income
        WriteTotalLine(table, 4);

        table.AcceptChanges();
        gridView.DataSource = table;
        gridView.DataBind();      

        //style adjust
        TableStyleAdjustment(table, gridView);
    }

    private void WriteHeaderByPlant(DataTable table)
    {
        //Header Columns: no duplicated column name allowed, 
        //so if use space, must increase the # of space to specify the columname.
        table.Columns.Add("Name");                          //column 1
        table.Columns.Add("Period " + fiscalMonth + " Actual");               //column 2      allsummary[0, j]
        table.Columns.Add(" ");                             //column 3      allsummary[1, j]
        table.Columns.Add("Last Period Actual");            //column 4      allsummary[2, j]
        table.Columns.Add("  ");                            //column 5      allsummary[3, j]
        table.Columns.Add("Period " + fiscalMonth + " Budget");               //column 6      allsummary[4, j]
        table.Columns.Add("   ");                           //column 7      allsummary[5, j]
        table.Columns.Add("Y-T-D Actual");                  //column 8      allsummary[6, j]
        table.Columns.Add("    ");                          //column 9      allsummary[7, j]
        table.Columns.Add("Y-T-D Budget");                  //column 10     allsummary[8, j]
        table.Columns.Add("     ");                         //column 11     allsummary[9, j]
        table.Columns.Add("Last Year Period " + fiscalMonth + " Actual");     //column 12     allsummary[10, j]
        table.Columns.Add("      ");                        //column 13     allsummary[11, j]
    }

    private string NegativeCurrencyFormat(double getAmount, bool Ignore = false)
    {
        if (getAmount < 0)
        {
            //return "-" + getAmount.ToString("C2");
            return getAmount.ToString("C2");
        }
        else
        {
            return getAmount.ToString("C2");
        }
    }

    private string ForcePositiveCurrency(double getAmount, bool Ignore = false)
    {
        if (getAmount >= 0)
        {
            //return "-" + getAmount.ToString("C2");
            return getAmount.ToString("C2");
        }
        else
        {
            return (getAmount * -1).ToString("C2");
        }
    }

    private bool isZero(double a)
    {
        if ((a > 0.000001) || (a < 0.000001))
            return false;
        else
            return true;
    }

    private void WriteAmountConsolidate(string blockname, DataTable table, List<IncomeStatementReport.Categories.Group> groupList, int block, ExcoCalendar calendar)
    {
        bool flagInsert = false;
        bool hasProductionSales = false;
        bool hasSteelSurcharge = false;
        bool hasOtherSales = false;

        double[] Column_Total = new double[15];
        double[] Surcharge_Total = new double[15];

        // set to 0
        for (int i = 0; i < 15; i++)
        {
            Column_Total[i] = 0;
        }

        DataRow blockrow = table.NewRow();
        blockrow["Name"] = blockname;
        table.Rows.Add(blockrow);
        blockposition[block, 0] = table.Rows.Count - 1;
        double curMa, curMi, curTe, curCo;

        if (blockname == "") provisionRow = table.Rows.Count;

        switchpositivenegative = ((blockname == "SALES") ? 1 : -1);

        bool is_Surcharge = false;

        foreach (IncomeStatementReport.Categories.Group group in groupList)
        {

            if (!hasProductionSales && group.name.Contains("CANADA SALES") && Implement_Changes)
            {
                //Add empty row
                //DataRow emprow2 = table.NewRow();
                //table.Rows.Add(emprow2);

                DataRow datarow2 = table.NewRow();
                datarow2[0] = "PRODUCTION SALES";
                hasProductionSales = true;

                table.Rows.Add(datarow2);
                line_subgroup[0] = table.Rows.Count - 1;
                blockposition[block, 0] += 2;

            }
            else if (!hasSteelSurcharge && group.name.Contains("CANADA SALES STEEL SURCHARGE") && Implement_Changes)
            {
                //Add empty row
                DataRow emprow2 = table.NewRow();

                DataRow datarow2 = table.NewRow();
                datarow2[0] = "TOTAL PRODUCTION SALES";
                hasProductionSales = true;

                is_Surcharge = true;

                for (int i = 0; i <= 10; i++)
                {
                    datarow2[i + 1] = NegativeCurrencyFormat(Column_Total[i]);
                    Column_Total[i] = 0; // reset
                }

                table.Rows.Add(datarow2);
                total_line_row_number[0] = table.Rows.Count - 1;
                table.Rows.Add(emprow2);

                //emprow2[0] = "STEEL SURCHARGE";
                //blockposition[block, 0] += 2;


                DataRow datarow3 = table.NewRow();
                datarow3[0] = "STEEL SURCHARGE";
                //table.Rows.Add(datarow3);
                //line_subgroup[1] = table.Rows.Count - 1;
                hasSteelSurcharge = true;


            }
            else if (!hasOtherSales && group.name.Contains("CASH & VOLUMN DISCOUNTS") && Implement_Changes)
            {
                //Add empty row
                DataRow emprow2 = table.NewRow();

                DataRow datarow2 = table.NewRow();
                datarow2[0] = "TOTAL STEEL SURCHARGE";
                hasProductionSales = true;


                for (int i = 0; i <= 10; i++)
                {
                    Surcharge_Total[i] = Column_Total[i];
                    Column_Total[i] = 0; // reset
                }

                //table.Rows.Add(datarow2);
                //total_line_row_number[1] = table.Rows.Count - 1;
                //table.Rows.Add(emprow2);

                //emprow2[0] = "STEEL SURCHARGE";
                //blockposition[block, 0] += 2;


                DataRow datarow3 = table.NewRow();
                datarow3[0] = "OTHER SALES";
                table.Rows.Add(datarow3);
                line_subgroup[2] = table.Rows.Count - 1;
                hasOtherSales = true;
                datarow2 = table.NewRow();
                datarow2[0] = "STEEL SURCHARGE";
                hasProductionSales = true;

                for (int i = 0; i <= 10; i++)
                {
                    datarow2[i + 1] = NegativeCurrencyFormat(Surcharge_Total[i]);
                    Column_Total[i] += Surcharge_Total[i];
                }

                is_Surcharge = false;
                table.Rows.Add(datarow2);
            }

            bool isProvision = false; 

            switch (group.name)
            {
                case "PROVISION FOR INCOME TAX":
                    isProvision = true;
                    provisionRow = table.Rows.Count;
                    break;
                case "FACTORY SOUTH AMERICA TRAVEL":
                    factorySATravelRow = table.Rows.Count;
                    break;
                case "OFFICE SOUTH AMERICA TRAVEL":
                    officeSATravelRow = table.Rows.Count;
                    break;
                case "DIRECT LABOUR SHOP":
                    directLabourShop = table.Rows.Count;
                    break;
                case "DIRECT LABOUR HEAT TREAT":
                    directLabourHeat = table.Rows.Count;
                    break;
                case "CAM SALARIES":
                    directLabourCAM = table.Rows.Count;
                    break;
                case "CAD SALARIES":
                    factoryCAD = table.Rows.Count;
                    break;
                case "INDIRECT LABOUR ISO IT PURCHASER":
                    factoryIndirectLabour = table.Rows.Count;
                    break;
                case "SUPERVISORY SALARIES":
                    factorySupervisor = table.Rows.Count;
                    break;
                case "FACTORY EMPLOYEE BENEFITS":
                    factoryEmployeeBenefits = table.Rows.Count;
                    break;
                case "FACTORY GROUP INSURANCE":
                    factoryGroupInsurance = table.Rows.Count;
                    break;
                case "STATUTORY HOLIDAY":
                    factoryStatHoliday = table.Rows.Count;
                    break;
                case "VACATION PAY EXPENSE":
                    factoryVacationPay = table.Rows.Count;
                    break;
                case "WORKERS COMPENSATION":
                    factoryWorkersCompensation = table.Rows.Count;
                    break;
                case "SALARIES":
                    officeSalaries = table.Rows.Count;
                    break;
                case "SALES SALARIES":
                    salesSalaries = table.Rows.Count;
                    break;
                case "EMPLOYEE BENEFITS SALES":
                    salesEmployeeBenefits = table.Rows.Count;
                    break; 
                case "GROUP INSURANCE SALES":
                    salesGroupInsurance = table.Rows.Count;
                    break;
                case "STATUTORY HOLIDAY SALES":
                    salesStatHoliday = table.Rows.Count;
                    break;
                case "VACATION PAY SALES":
                    salesVacationPay = table.Rows.Count;
                    break;
                case "WORKERS COMPENSATION SALES":
                    salesWorkersCompensation = table.Rows.Count;
                    break;
                case "OFFICE EMPLOYEE BENEFITS":
                    officeEmployeeBenefits = table.Rows.Count;
                    break;
                case "OFFICE GROUP INSURANCE":
                    officeGroupInsurance = table.Rows.Count;
                    break;
                case "PAYROLL TAXES":
                    officePayrollTaxes = table.Rows.Count;
                    break;
                case "OFFICE STATUTORY HOLIDAY":
                    officeStatHoliday = table.Rows.Count;
                    break;
                case "OFFICE VACATION PAY":
                    officeVacationPay = table.Rows.Count;
                    break;
                case "OFFICE WORKER'S COMPENSATION":
                    officeWorkersCompensation = table.Rows.Count;
                    break;
                case "OFFICE RENT":
                    officeRent = table.Rows.Count;
                    break;
                case "FACTORY RENT":
                    factoryRent = table.Rows.Count;
                    break;
                case "DEPRECIATION BUILDING":
                    factoryDepreBuilding = table.Rows.Count;
                    break;
                case "DEPRECIATION BUILDING IMPROVEMENT":
                    factoryDepreBuildingImprov = table.Rows.Count;
                    break;
                case "MAINTENANCE AND REPAIR BUILDING":
                    factoryReqMainBuilding = table.Rows.Count;
                    break;
                case "REALTY TAX":
                    factoryPropertyTax = table.Rows.Count;
                    break;
                case "DEPRECIATION MACHINE AND EQUIPMENT":
                    factoryDepreMachine = table.Rows.Count;
                    break;
                case "DEPRECIATION FURNITURE AND FIXTURE":
                    factoryDepreFurniture = table.Rows.Count;
                    break;
                case "DEPRECIATION SOFTWARE":
                    factoryDepreSoftware = table.Rows.Count;
                    break;
                case "OFFICE DEPRECIATION EXPENSE":
                    officeDepreFurniture = table.Rows.Count;
                    break;
                case "SOFTWARE AMORTIZATION":
                    officeSoftwareAmort = table.Rows.Count;
                    break;
                case "TOOL AMORTIZATION":
                    factoryToolAmort = table.Rows.Count;
                    break;
                case "DEPRECIATION AUTOMOBILE":
                    salesDepreAuto = table.Rows.Count;
                    break;
                case "TOOLS EXPENSE":
                    factoryToolExpense = table.Rows.Count;
                    break;
                case "SHIPPING SUPPLIES":
                    factoryShippingSupplies = table.Rows.Count;
                    break;
                case "SHOP SUPPLIES EXPENSE":
                    factoryShopSupplies = table.Rows.Count;
                    break;
                case "SHOP SUPPLIES HEAT TREAT":
                    factoryHeatTreatSupplies = table.Rows.Count;
                    break;
                case "CAD and CAM SUPPLIES":
                    factoryCADCAMSupplies = table.Rows.Count;
                    break;
                case "OFFICE SUPPLIES":
                    officeSupplies = table.Rows.Count;
                    break;
                case "SHOP TRAVEL":
                    factoryTravel = table.Rows.Count;
                    break;
                case "SHOP MEALS AND ENTERTAINMENT":
                    factoryMeals = table.Rows.Count;
                    break;
                case "SHOP AIR FARE":
                    factoryAirFare = table.Rows.Count;
                    break;
                case "NON DEDUCTIBLE EXPENSE GOLF":
                    salesGolf = table.Rows.Count;
                    break;
                case "SELLING AND TRAVEL MEALS AND ENTERTAINMENT":
                    salesMeals = table.Rows.Count;
                    break;
                case "SELLING AND TRAVEL EXPENSES":
                    salesTravel = table.Rows.Count;
                    break;
                case "AIR FARE":
                    salesAirFare = table.Rows.Count;
                    break;
                case "MEAL AND ENTERTAINMENT":
                    officeMeals = table.Rows.Count;
                    break;
                case "TRAVEL EXPENSES":
                    officeTravel = table.Rows.Count;
                    break;
                case "OFFICE AIR FARE":
                    officeAirFare = table.Rows.Count;
                    break;
            }
            curMa = 0.00;
            curMi = 0.00;
            curTe = 0.00;

            curCo = 0.00;
            DataRow datarow = table.NewRow();
            //column 1
            datarow[0] = group.name;

            
            //column 2
            datarow[1] = NegativeCurrencyFormat((group.plant01.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative));
            allsummary[block, 0] += (group.plant01.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative);
            Column_Total[0] += (group.plant01.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative);
            if (blockname.Length == 0) provisionalTax[1] = (group.plant01.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative);
            curMa = (group.plant01.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative);  
            //column 4
            datarow[3] = NegativeCurrencyFormat((group.plant03.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative));
            allsummary[block, 2] += (group.plant03.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative);
            Column_Total[2] += (group.plant03.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative);
            if (blockname.Length == 0) provisionalTax[3] =  (group.plant03.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative);
            curMi = (group.plant03.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative);
            //column 6
            datarow[5] = NegativeCurrencyFormat((group.plant05.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative));
            allsummary[block, 4] += (group.plant05.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative);
            Column_Total[4] += (group.plant05.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative);
            if (blockname.Length == 0) provisionalTax[5] =  (group.plant05.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative);
            curTe = (group.plant05.actualThisYearList[fiscalMonth].GetAmount("CA") * switchpositivenegative);
            //column 8
            if (fiscalMonth == 1)
            {
                datarow[7] = NegativeCurrencyFormat((group.plant04.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant41.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant48.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant49.actualThisYearList[fiscalMonth].GetAmount("CA")) * switchpositivenegative);
                allsummary[block, 6] += ((group.plant04.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant41.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant48.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant49.actualThisYearList[fiscalMonth].GetAmount("CA")) * switchpositivenegative);
                Column_Total[6] += ((group.plant04.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant41.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant48.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant49.actualThisYearList[fiscalMonth].GetAmount("CA")) * switchpositivenegative);
                if (blockname.Length == 0) provisionalTax[7] =  ((group.plant04.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant41.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant48.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant49.actualThisYearList[fiscalMonth].GetAmount("CA")) * switchpositivenegative);
                curCo = ((group.plant04.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant41.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant48.actualThisYearList[fiscalMonth].GetAmount("CA") + group.plant49.actualThisYearList[fiscalMonth].GetAmount("CA")) * switchpositivenegative);
            }
            else
            {
                // ROBIN HERE
                int g = group.plant04.calendar.GetCalendarYear();

                ExcoCalendar calendarx = new ExcoCalendar(g == 0 ? 17 : g, group.plant04.calendar.GetCalendarMonth(), true, 1);
                datarow[7] = NegativeCurrencyFormat(((new ExcoMoney(calendarx.GetNextCalendarMonth(), (group.plant04.actualThisYearList[fiscalMonth] + group.plant41.actualThisYearList[fiscalMonth] + group.plant48.actualThisYearList[fiscalMonth] + group.plant49.actualThisYearList[fiscalMonth]).amountCP - (group.plant04.actualThisYearList[fiscalMonth - 1] + group.plant41.actualThisYearList[fiscalMonth - 1] + group.plant48.actualThisYearList[fiscalMonth - 1] + group.plant49.actualThisYearList[fiscalMonth - 1]).amountCP, "CP")).GetAmount("CA")) * switchpositivenegative);
                allsummary[block, 6] += ((new ExcoMoney(calendarx.GetNextCalendarMonth(), (group.plant04.actualThisYearList[fiscalMonth] + group.plant41.actualThisYearList[fiscalMonth] + group.plant48.actualThisYearList[fiscalMonth] + group.plant49.actualThisYearList[fiscalMonth]).amountCP - (group.plant04.actualThisYearList[fiscalMonth - 1] + group.plant41.actualThisYearList[fiscalMonth - 1] + group.plant48.actualThisYearList[fiscalMonth - 1] + group.plant49.actualThisYearList[fiscalMonth - 1]).amountCP, "CP")).GetAmount("CA") * switchpositivenegative);
                Column_Total[6] += ((new ExcoMoney(calendarx.GetNextCalendarMonth(), (group.plant04.actualThisYearList[fiscalMonth] + group.plant41.actualThisYearList[fiscalMonth] + group.plant48.actualThisYearList[fiscalMonth] + group.plant49.actualThisYearList[fiscalMonth]).amountCP - (group.plant04.actualThisYearList[fiscalMonth - 1] + group.plant41.actualThisYearList[fiscalMonth - 1] + group.plant48.actualThisYearList[fiscalMonth - 1] + group.plant49.actualThisYearList[fiscalMonth - 1]).amountCP, "CP")).GetAmount("CA") * switchpositivenegative);
                if (blockname.Length == 0) provisionalTax[7] =  ((new ExcoMoney(calendarx.GetNextCalendarMonth(), (group.plant04.actualThisYearList[fiscalMonth] + group.plant41.actualThisYearList[fiscalMonth] + group.plant48.actualThisYearList[fiscalMonth] + group.plant49.actualThisYearList[fiscalMonth]).amountCP - (group.plant04.actualThisYearList[fiscalMonth - 1] + group.plant41.actualThisYearList[fiscalMonth - 1] + group.plant48.actualThisYearList[fiscalMonth - 1] + group.plant49.actualThisYearList[fiscalMonth - 1]).amountCP, "CP")).GetAmount("CA") * switchpositivenegative);
                curCo = ((new ExcoMoney(calendarx.GetNextCalendarMonth(), (group.plant04.actualThisYearList[fiscalMonth] + group.plant41.actualThisYearList[fiscalMonth] + group.plant48.actualThisYearList[fiscalMonth] + group.plant49.actualThisYearList[fiscalMonth]).amountCP - (group.plant04.actualThisYearList[fiscalMonth - 1] + group.plant41.actualThisYearList[fiscalMonth - 1] + group.plant48.actualThisYearList[fiscalMonth - 1] + group.plant49.actualThisYearList[fiscalMonth - 1]).amountCP, "CP")).GetAmount("CA") * switchpositivenegative);
            }            
            //column 10
            datarow[9] = NegativeCurrencyFormat(curMa + curMi + curTe + curCo);
            allsummary[block, 8] += curMa + curMi + curTe + curCo;
            Column_Total[8] += curMa + curMi + curTe + curCo;
            if (blockname.Length == 0) provisionalTax[9] = curMa + curMi + curTe + curCo;

            double ytdTotal = 0;
            //column 12
            for (int i = 1; i < fiscalMonth + 1; i++)
            {
                int g = group.plant04.calendar.GetCalendarYear();
                ExcoCalendar calendarx = new ExcoCalendar(g == 0 ? 17 : g, group.plant04.calendar.GetCalendarMonth(), true, 1);
                ytdTotal += ((new ExcoMoney(calendarx.GetNextCalendarMonth(), (group.plant04.actualThisYearList[i] + group.plant41.actualThisYearList[i] + group.plant48.actualThisYearList[i] + group.plant49.actualThisYearList[i]).amountCP - (group.plant04.actualThisYearList[i - 1] + group.plant41.actualThisYearList[i - 1] + group.plant48.actualThisYearList[i - 1] + group.plant49.actualThisYearList[i - 1]).amountCP, "CP")).GetAmount("CA")) * switchpositivenegative;
                
                ytdTotal += group.plant01.actualThisYearList[i].GetAmount("CA") * switchpositivenegative;//..Sum(g => g.GetAmount("CA"));
                ytdTotal += group.plant03.actualThisYearList[i].GetAmount("CA") * switchpositivenegative;//.Sum(g => g.GetAmount("CA"));
                //ytdTotal += (group.plant04.actualThisYearList[i].GetAmount("CA") + group.plant41.actualThisYearList[i].GetAmount("CA") + group.plant48.actualThisYearList[i].GetAmount("CA") + group.plant49.actualThisYearList[i].GetAmount("CA")) * switchpositivenegative;//.Sum(g => g.GetAmount("CA"));
                ytdTotal += group.plant05.actualThisYearList[i].GetAmount("CA") * switchpositivenegative;//.Sum(g => g.GetAmount("CA")); 
            }
            if (blockname.Length == 0) provisionalTax[11] += ytdTotal;

            //ytdTotal *= ytdTotal < 0 ? -1 : 1;
            datarow[11] = NegativeCurrencyFormat(ytdTotal);
            allsummary[block, 10] += ytdTotal;
            Column_Total[10] += ytdTotal;

            if (!is_Surcharge) table.Rows.Add(datarow);

            if (group.name.Contains("SCRAP SALES"))
            {
                flagInsert = true;
            }
        }

        if (flagInsert && Implement_Changes)
        {
            //Add empty row
            DataRow emprow2 = table.NewRow();

            DataRow datarow2 = table.NewRow();
            datarow2[0] = "TOTAL OTHER SALES";
            hasProductionSales = true;


            for (int i = 0; i <= 10; i++)
            {
                datarow2[i + 1] = NegativeCurrencyFormat(Column_Total[i]);
                Column_Total[i] = 0; // reset
            }

            table.Rows.Add(datarow2);
            total_line_row_number[2] = table.Rows.Count - 1;
            table.Rows.Add(emprow2);

        }

        if (blockname.Length > 0)
        {
            DataRow total = table.NewRow();
            total[0] = "TOTAL:";
            total[1] = NegativeCurrencyFormat(allsummary[block, 0]);
            total[3] = NegativeCurrencyFormat(allsummary[block, 2]);
            total[5] = NegativeCurrencyFormat(allsummary[block, 4]);
            total[7] = NegativeCurrencyFormat(allsummary[block, 6]);
            total[9] = NegativeCurrencyFormat(allsummary[block, 8]);
            total[11] = NegativeCurrencyFormat(allsummary[block, 10]);
            table.Rows.Add(total);

            blockposition[block, 1] = table.Rows.Count - 1;
        }
        else if (block == 7) 
        {

            blockposition[block, 1] = table.Rows.Count;
        }


        //column 3,5,7,9,11
        for (int i = blockposition[block, 0]; i < blockposition[block, 1]; i++)
        {

            if (isZero(allsummary[0, 0]) == false && ((table.Rows[i][1].ToString().Length > 0) || false))
            {
                int p = 1;
                tempdata = table.Rows[i][1].ToString();
                tempp = double.Parse(tempdata, styles);
                table.Rows[i][2] = (tempp / allsummary[0, 0]).ToString("P2");
                allsummary[block, 1] += tempp / allsummary[0, 0];
                Column_Total[1] += tempp / allsummary[0, 0];
                if (blockname.Length == 0) provisionalTax[2] =  tempp / allsummary[0, 0];
            }
            else if (table.Rows[i][1].ToString().Length > 0 || false)
            {
                table.Rows[i][2] = "0.00%";
                allsummary[block, 1] = 0.00;
                Column_Total[1] = 0.00;
                if (blockname.Length == 0) provisionalTax[2] =  0.00;
            }

            if (isZero(allsummary[0, 2]) == false && ((table.Rows[i][3].ToString().Length > 0) || false))
            {
                tempdata = table.Rows[i][3].ToString();
                tempp = double.Parse(tempdata, styles);
                table.Rows[i][4] = (tempp / allsummary[0, 2]).ToString("P2");
                allsummary[block, 3] += tempp / allsummary[0, 2];
                Column_Total[3] += tempp / allsummary[0, 2];
                if (blockname.Length == 0) provisionalTax[4] =  tempp / allsummary[0, 2];
            }
            else if (table.Rows[i][3].ToString().Length > 0 || false)
            {
                table.Rows[i][4] = "0.00%";
                allsummary[block, 3] = 0.00;
                Column_Total[3] = 0.00; 
                if (blockname.Length == 0) provisionalTax[4] = 0.00;
            }

            if (isZero(allsummary[0, 4]) == false && ((table.Rows[i][5].ToString().Length > 0) || false))
            {
                tempdata = table.Rows[i][5].ToString();
                tempp = double.Parse(tempdata, styles);
                table.Rows[i][6] = (tempp / allsummary[0, 4]).ToString("P2");
                allsummary[block, 5] += tempp / allsummary[0, 4];
                Column_Total[5] += tempp / allsummary[0, 4];
                if (blockname.Length == 0) provisionalTax[6] = tempp / allsummary[0, 4];
            }
            else if (table.Rows[i][5].ToString().Length > 0 || false)
            {
                table.Rows[i][6] = "0.00%";
                allsummary[block, 5] = 0.00;
                Column_Total[5] = 0.00;
                if (blockname.Length == 0) provisionalTax[6] = 0.00;
            }

            if (isZero(allsummary[0, 6]) == false && ((table.Rows[i][7].ToString().Length > 0) || false))
            {
                tempdata = table.Rows[i][7].ToString();
                tempp = double.Parse(tempdata, styles);
                table.Rows[i][8] = (tempp / allsummary[0, 6]).ToString("P2");
                allsummary[block, 7] += tempp / allsummary[0, 6];
                Column_Total[7] += tempp / allsummary[0, 6];
                if (blockname.Length == 0) provisionalTax[8] =  tempp / allsummary[0, 6];
            }
            else if (table.Rows[i][7].ToString().Length > 0 || false)
            {
                table.Rows[i][8] = "0.00%";
                allsummary[block, 7] = 0.00;
                Column_Total[7] = 0.00;
                if (blockname.Length == 0) provisionalTax[8] = 0.00;
            }

            if (isZero(allsummary[0, 8]) == false && ((table.Rows[i][9].ToString().Length > 0) || false))
            {
                tempdata = table.Rows[i][9].ToString();
                tempp = double.Parse(tempdata, styles);
                table.Rows[i][10] = (tempp / allsummary[0, 8]).ToString("P2");
                allsummary[block, 9] += tempp / allsummary[0, 8];
                Column_Total[9] += tempp / allsummary[0, 8];
                if (blockname.Length == 0) provisionalTax[10] =  tempp / allsummary[0, 8];
            }
            else if (table.Rows[i][9].ToString().Length > 0 || false)
            {
                table.Rows[i][10] = "0.00%";
                allsummary[block, 9] = 0.00;
                Column_Total[9] = 0.00;
                if (blockname.Length == 0) provisionalTax[10] = 0.00;
            }

        }
        
        if (blockname.Length > 0)
        {
            table.Rows[blockposition[block, 1]][2] = (allsummary[block, 1] > 1 ? allsummary[block, 1] - 1 : allsummary[block, 1]).ToString("P2");
            table.Rows[blockposition[block, 1]][4] = (allsummary[block, 3] > 1 ? allsummary[block, 3] - 1 : allsummary[block, 3]).ToString("P2");
            table.Rows[blockposition[block, 1]][6] =( allsummary[block, 5] > 1 ? allsummary[block, 5] - 1 : allsummary[block, 5]).ToString("P2");
            table.Rows[blockposition[block, 1]][8] = (allsummary[block, 7] > 1 ? allsummary[block, 7] - 1 : allsummary[block, 7]).ToString("P2");
            table.Rows[blockposition[block, 1]][10] = (allsummary[block, 9] > 1 ? allsummary[block, 9] - 1 : allsummary[block, 9]).ToString("P2");
        }
        else
        {
            //table.Rows[blockposition[block, 1]][2] = provisionalTax[2].ToString("P2");
            //table.Rows[blockposition[block, 1]][4] = provisionalTax[4].ToString("P2");
            //table.Rows[blockposition[block, 1]][6] = provisionalTax[6].ToString("P2");
            //table.Rows[blockposition[block, 1]][8] = provisionalTax[8].ToString("P2");
            //table.Rows[blockposition[block, 1]][10] = provisionalTax[10].ToString("P2");
        }

        //used to separate each block
        DataRow emprow = table.NewRow();
        table.Rows.Add(emprow);  
    }



    private void WriteAmount(string blockname, DataTable table, int plantID, List<IncomeStatementReport.Categories.Group> groupList, int block, string currency)
    {
        bool flagInsert = false;
        bool hasProductionSales = false;
        bool hasSteelSurcharge = false;
        bool hasOtherSales = false;

        DataRow blockrow = table.NewRow();
        double[] Column_Total = new double[15];
        double[] Surcharge_Row = new double[15];
        bool is_Surcharge = false;

        // set to 0
        for (int i = 0; i < 15; i++)
        {
            Column_Total[i] = 0;
        }

        if (blockname.Length > 0)
        {
            blockrow["Name"] = blockname;
            table.Rows.Add(blockrow);
        }

        blockposition[block, 0] = table.Rows.Count - 1;
        foreach (IncomeStatementReport.Categories.Group group in groupList)
        {
            
            if (!hasProductionSales && group.name.Contains("CANADA SALES") && Implement_Changes)
            {
                //Add empty row
                //DataRow emprow2 = table.NewRow();
                //table.Rows.Add(emprow2);  

                DataRow datarow2 = table.NewRow();
                datarow2[0] = "PRODUCTION SALES";
                hasProductionSales = true;

                table.Rows.Add(datarow2);
                line_subgroup[0] = table.Rows.Count - 1;
                blockposition[block, 0] += 2;
            }
            else if (!hasSteelSurcharge && group.name.Contains("CANADA SALES STEEL SURCHARGE") && Implement_Changes)
            {
                //Add empty row
                DataRow emprow2 = table.NewRow();

                DataRow datarow2 = table.NewRow();
                datarow2[0] = "TOTAL PRODUCTION SALES";
                hasProductionSales = true;


                for (int i = 0; i <= 11; i++)
                {
                    datarow2[i + 1] = NegativeCurrencyFormat(Column_Total[i]);
                    Column_Total[i] = 0; // reset
                }

                table.Rows.Add(datarow2);
                total_line_row_number[0] = table.Rows.Count - 1;
                table.Rows.Add(emprow2);

                //emprow2[0] = "STEEL SURCHARGE";
                //blockposition[block, 0] += 2;

                is_Surcharge = true;
                //DataRow datarow3 = table.NewRow();
                //datarow3[0] = "STEEL SURCHARGE";
                //table.Rows.Add(datarow3);
                //line_subgroup[1] = table.Rows.Count - 1;
                hasSteelSurcharge = true;
            }
            else if (!hasOtherSales && group.name.Contains("CASH & VOLUMN DISCOUNTS") && Implement_Changes)
            {
                //Add empty row
                //DataRow emprow2 = table.NewRow();

                //DataRow datarow2 = table.NewRow();
                //datarow2[0] = "TOTAL STEEL SURCHARGE";
                //hasProductionSales = true;


                for (int i = 0; i <= 11; i++)
                {
                    Surcharge_Row[i] = Column_Total[i]; // assign total surcharge to array
                    Column_Total[i] = 0; // reset
                }

                //table.Rows.Add(datarow2);
                //total_line_row_number[1] = table.Rows.Count - 1;
                //table.Rows.Add(emprow2);

                //emprow2[0] = "STEEL SURCHARGE";
                //blockposition[block, 0] += 2;

                
                DataRow datarow3 = table.NewRow();
                datarow3[0] = "OTHER SALES";
                table.Rows.Add(datarow3);
                line_subgroup[2] = table.Rows.Count - 1;
                hasOtherSales = true;
                is_Surcharge = false; 
                
                DataRow datarow2 = table.NewRow();
                datarow2[0] = "STEEL SURCHARGE";
                hasProductionSales = true;

                for (int i = 0; i <= 11; i++)
                {
                    datarow2[i + 1] = NegativeCurrencyFormat(Surcharge_Row[i]);
                }

                table.Rows.Add(datarow2);
            }
                /*
            else if (flagInsert && Implement_Changes)
            {
                //Add empty row
                DataRow emprow2 = table.NewRow();

                DataRow datarow2 = table.NewRow();
                datarow2[0] = "TOTAL OTHER SALES";
                hasProductionSales = true;


                for (int i = 0; i <= 11; i++)
                {
                    datarow2[i + 1] = NegativeCurrencyFormat(Column_Total[i]);
                    Column_Total[i] = 0; // reset
                }

                table.Rows.Add(datarow2);
                total_line_row_number[0] = table.Rows.Count - 1;
                table.Rows.Add(emprow2);

            }*/
            if (group.name.Contains("SCRAP SALES"))
            {
                flagInsert = true;
            }

            DataRow datarow = table.NewRow();
            //column 1
            datarow[0] = group.name;
            
            Plant plant;
            if (plantID == 5)
            {
                plant = group.plant05;
            }
            else if (plantID == 1)
            {
                plant = group.plant01;
            }
            else if (plantID == 3)
            {
                plant = group.plant03;
            }
            else
            {
                plant = group.plant04;
            }

            if (4 != plantID)
            {
                //column 2
                datarow[1] = NegativeCurrencyFormat((plant.actualThisYearList[fiscalMonth].GetAmount(currency) * switchpositivenegative));
                allsummary[block, 0] += (plant.actualThisYearList[fiscalMonth].GetAmount(currency) * switchpositivenegative);
                Column_Total[0] += (plant.actualThisYearList[fiscalMonth].GetAmount(currency) * switchpositivenegative);
                if (blockname.Length == 0) provisionalTax[1] = (plant.actualThisYearList[fiscalMonth].GetAmount(currency) * switchpositivenegative);
                //column 4
                if (fiscalMonth > 1)
                {
                    datarow[3] = NegativeCurrencyFormat((plant.actualThisYearList[fiscalMonth - 1].GetAmount(currency) * switchpositivenegative));
                    allsummary[block, 2] += (plant.actualThisYearList[fiscalMonth - 1].GetAmount(currency)*switchpositivenegative);
                    Column_Total[2] += (plant.actualThisYearList[fiscalMonth - 1].GetAmount(currency)*switchpositivenegative);
                    if (blockname.Length == 0) provisionalTax[3] = (plant.actualThisYearList[fiscalMonth - 1].GetAmount(currency) * switchpositivenegative);
                }
                else
                {
                    datarow[3] = NegativeCurrencyFormat((plant.actualLastYearList[12].GetAmount(currency) * switchpositivenegative));
                    //allsummary[block, 2] += (plant.actualThisYearList[fiscalMonth - 1].GetAmount(currency)*switchpositivenegative); //tiger error
                    allsummary[block, 2] += (plant.actualLastYearList[12].GetAmount(currency) * switchpositivenegative);
                    Column_Total[2] += (plant.actualLastYearList[12].GetAmount(currency) * switchpositivenegative);
                    if (blockname.Length == 0) provisionalTax[3] = (plant.actualLastYearList[12].GetAmount(currency) * switchpositivenegative);
                }
                //column 6
                datarow[5] = NegativeCurrencyFormat((plant.budgetThisYearList[fiscalMonth].GetAmount(currency) * switchpositivenegative));
                allsummary[block, 4] += (plant.budgetThisYearList[fiscalMonth].GetAmount(currency) * switchpositivenegative);
                Column_Total[4] += (plant.budgetThisYearList[fiscalMonth].GetAmount(currency) * switchpositivenegative);
                if (blockname.Length == 0) provisionalTax[5] = (plant.budgetThisYearList[fiscalMonth].GetAmount(currency) * switchpositivenegative);
                //column 8 ,10
                ytdActual = 0.0;
                ytdBudget = 0.0;
                for (int i = 1; i <= fiscalMonth; i++)
                {
                    ytdActual += plant.actualThisYearList[i].GetAmount(currency);
                    ytdBudget += plant.budgetThisYearList[i].GetAmount(currency);
                }
                datarow[7] = NegativeCurrencyFormat((ytdActual*switchpositivenegative));
                allsummary[block, 6] += (ytdActual*switchpositivenegative);
                Column_Total[6] += (ytdActual * switchpositivenegative);
                if (blockname.Length == 0) provisionalTax[7] = (ytdActual * switchpositivenegative);
                datarow[9] = NegativeCurrencyFormat((ytdBudget*switchpositivenegative));
                allsummary[block, 8] += (ytdBudget*switchpositivenegative);
                Column_Total[8] += (ytdBudget * switchpositivenegative);
                if (blockname.Length == 0) provisionalTax[9] = (ytdBudget * switchpositivenegative);
                //column 12
                datarow[11] = NegativeCurrencyFormat((plant.actualLastYearList[fiscalMonth].GetAmount(currency)*switchpositivenegative)); 
                allsummary[block, 10] += (plant.actualLastYearList[fiscalMonth].GetAmount(currency)*switchpositivenegative);
                Column_Total[10] += (plant.actualLastYearList[fiscalMonth].GetAmount(currency) * switchpositivenegative);
                if (blockname.Length == 0) provisionalTax[11] = (plant.actualLastYearList[fiscalMonth].GetAmount(currency) * switchpositivenegative);

                if (!is_Surcharge) table.Rows.Add(datarow);
            }
            else
            {
                ExcoMoney thisPeriod = new ExcoMoney();
                ExcoMoney lastPeriod = new ExcoMoney();
                ExcoMoney thisPeriodLastYear = new ExcoMoney();
                ExcoCalendar calendar = new ExcoCalendar(group.plant04.calendar.GetCalendarYear(), group.plant04.calendar.GetCalendarMonth(), true, 1);

                try
                {
                    if (fiscalMonth > 2)
                    {
                        thisPeriod = new ExcoMoney(calendar.GetNextCalendarMonth(), (group.plant04.actualThisYearList[fiscalMonth] + group.plant41.actualThisYearList[fiscalMonth] + group.plant48.actualThisYearList[fiscalMonth] + group.plant49.actualThisYearList[fiscalMonth]).amountCP - (group.plant04.actualThisYearList[fiscalMonth - 1] + group.plant41.actualThisYearList[fiscalMonth - 1] + group.plant48.actualThisYearList[fiscalMonth - 1] + group.plant49.actualThisYearList[fiscalMonth - 1]).amountCP, "CP");
                        lastPeriod = new ExcoMoney(calendar.GetLastCalendarMonth(), (group.plant04.actualThisYearList[fiscalMonth - 1] + group.plant41.actualThisYearList[fiscalMonth - 1] + group.plant48.actualThisYearList[fiscalMonth - 1] + group.plant49.actualThisYearList[fiscalMonth - 1]).amountCP - (group.plant04.actualThisYearList[fiscalMonth - 2] + group.plant41.actualThisYearList[fiscalMonth - 2] + group.plant48.actualThisYearList[fiscalMonth - 2] + group.plant49.actualThisYearList[fiscalMonth - 2]).amountCP, "CP");
                        thisPeriodLastYear = new ExcoMoney(calendar.GetCalendarMonthLastYear(), (group.plant04.actualLastYearList[fiscalMonth] + group.plant41.actualLastYearList[fiscalMonth] + group.plant48.actualLastYearList[fiscalMonth] + group.plant49.actualLastYearList[fiscalMonth]).amountCP - (group.plant04.actualLastYearList[fiscalMonth - 1] + group.plant41.actualLastYearList[fiscalMonth - 1] + group.plant48.actualLastYearList[fiscalMonth - 1] + group.plant49.actualLastYearList[fiscalMonth - 1]).amountCP, "CP");
                    }
                    else if (fiscalMonth == 2)
                    {
                        thisPeriod = new ExcoMoney(calendar.GetNextCalendarMonth(), (group.plant04.actualThisYearList[fiscalMonth] + group.plant41.actualThisYearList[fiscalMonth] + group.plant48.actualThisYearList[fiscalMonth] + group.plant49.actualThisYearList[fiscalMonth]).amountCP - (group.plant04.actualThisYearList[fiscalMonth - 1] + group.plant41.actualThisYearList[fiscalMonth - 1] + group.plant48.actualThisYearList[fiscalMonth - 1] + group.plant49.actualThisYearList[fiscalMonth - 1]).amountCP, "CP");
                        lastPeriod = group.plant04.actualThisYearList[fiscalMonth - 1] + group.plant41.actualThisYearList[fiscalMonth - 1] + group.plant48.actualThisYearList[fiscalMonth - 1] + group.plant49.actualThisYearList[fiscalMonth - 1];
                        thisPeriodLastYear = new ExcoMoney(calendar.GetCalendarMonthLastYear(), (group.plant04.actualLastYearList[fiscalMonth] + group.plant41.actualLastYearList[fiscalMonth] + group.plant48.actualLastYearList[fiscalMonth] + group.plant49.actualLastYearList[fiscalMonth]).amountCP - (group.plant04.actualLastYearList[fiscalMonth - 1] + group.plant41.actualLastYearList[fiscalMonth - 1] + group.plant48.actualLastYearList[fiscalMonth - 1] + group.plant49.actualLastYearList[fiscalMonth - 1]).amountCP, "CP");
                    }
                    else
                    {
                        thisPeriod = group.plant04.actualThisYearList[fiscalMonth] + group.plant41.actualThisYearList[fiscalMonth] + group.plant48.actualThisYearList[fiscalMonth] + group.plant49.actualThisYearList[fiscalMonth];
                        lastPeriod = new ExcoMoney(calendar.GetLastCalendarMonth(), (group.plant04.actualLastYearList[12] + group.plant41.actualLastYearList[12] + group.plant48.actualLastYearList[12] + group.plant49.actualLastYearList[12]).amountCP - (group.plant04.actualLastYearList[11] + group.plant41.actualLastYearList[11] + group.plant48.actualLastYearList[11] + group.plant49.actualLastYearList[11]).amountCP, "CP");
                        thisPeriodLastYear = group.plant04.actualLastYearList[fiscalMonth] + group.plant41.actualLastYearList[fiscalMonth] + group.plant48.actualLastYearList[fiscalMonth] + group.plant49.actualLastYearList[fiscalMonth];
                    }
                }
                catch
                {
                }
                //column 2
                datarow[1] = NegativeCurrencyFormat((thisPeriod.GetAmount(currency) * switchpositivenegative));
                allsummary[block, 0] += (thisPeriod.GetAmount(currency) * switchpositivenegative);
                Column_Total[0] += (thisPeriod.GetAmount(currency) * switchpositivenegative);
                if (blockname.Length == 0) provisionalTax[1] = (thisPeriod.GetAmount(currency) * switchpositivenegative);
                //column 4
                datarow[3] = NegativeCurrencyFormat((lastPeriod.GetAmount(currency) * switchpositivenegative));
                allsummary[block, 2] += (lastPeriod.GetAmount(currency) * switchpositivenegative);
                Column_Total[2] += (lastPeriod.GetAmount(currency) * switchpositivenegative);
                if (blockname.Length == 0) provisionalTax[3] = (lastPeriod.GetAmount(currency) * switchpositivenegative);

                //column 6
                datarow[5] = NegativeCurrencyFormat((group.plant04.budgetThisYearList[fiscalMonth].GetAmount(currency) + group.plant41.budgetThisYearList[fiscalMonth].GetAmount(currency) + group.plant48.budgetThisYearList[fiscalMonth].GetAmount(currency) + group.plant49.budgetThisYearList[fiscalMonth].GetAmount(currency)) * switchpositivenegative);
                allsummary[block, 4] += (group.plant04.budgetThisYearList[fiscalMonth].GetAmount(currency) + group.plant41.budgetThisYearList[fiscalMonth].GetAmount(currency) + group.plant48.budgetThisYearList[fiscalMonth].GetAmount(currency) + group.plant49.budgetThisYearList[fiscalMonth].GetAmount(currency)) * switchpositivenegative;
                Column_Total[4] += (group.plant04.budgetThisYearList[fiscalMonth].GetAmount(currency) + group.plant41.budgetThisYearList[fiscalMonth].GetAmount(currency) + group.plant48.budgetThisYearList[fiscalMonth].GetAmount(currency) + group.plant49.budgetThisYearList[fiscalMonth].GetAmount(currency)) * switchpositivenegative;
                if (blockname.Length == 0) provisionalTax[5] = (group.plant04.budgetThisYearList[fiscalMonth].GetAmount(currency) + group.plant41.budgetThisYearList[fiscalMonth].GetAmount(currency) + group.plant48.budgetThisYearList[fiscalMonth].GetAmount(currency) + group.plant49.budgetThisYearList[fiscalMonth].GetAmount(currency)) * switchpositivenegative;
                //column 8 ,10
                ytdActual = 0.0;
                ytdBudget = 0.0;
                for (int i = 1; i <= fiscalMonth; i++)
                {
                    ytdBudget += group.plant04.budgetThisYearList[i].GetAmount(currency) + group.plant41.budgetThisYearList[i].GetAmount(currency) + group.plant48.budgetThisYearList[i].GetAmount(currency) + group.plant49.budgetThisYearList[i].GetAmount(currency);                
                }
                datarow[7] = NegativeCurrencyFormat(((group.plant04.actualThisYearList[fiscalMonth] + group.plant41.actualThisYearList[fiscalMonth] + group.plant48.actualThisYearList[fiscalMonth] + group.plant49.actualThisYearList[fiscalMonth]).GetAmount(currency) * switchpositivenegative));
                allsummary[block, 6] += ((group.plant04.actualThisYearList[fiscalMonth] + group.plant41.actualThisYearList[fiscalMonth] + group.plant48.actualThisYearList[fiscalMonth] + group.plant49.actualThisYearList[fiscalMonth]).GetAmount(currency) * switchpositivenegative);
                Column_Total[6] += ((group.plant04.actualThisYearList[fiscalMonth] + group.plant41.actualThisYearList[fiscalMonth] + group.plant48.actualThisYearList[fiscalMonth] + group.plant49.actualThisYearList[fiscalMonth]).GetAmount(currency) * switchpositivenegative);
                if (blockname.Length == 0) provisionalTax[7] = ((group.plant04.actualThisYearList[fiscalMonth] + group.plant41.actualThisYearList[fiscalMonth] + group.plant48.actualThisYearList[fiscalMonth] + group.plant49.actualThisYearList[fiscalMonth]).GetAmount(currency) * switchpositivenegative);
                datarow[9] = NegativeCurrencyFormat((ytdBudget*switchpositivenegative));
                allsummary[block, 8] += (ytdBudget*switchpositivenegative);
                Column_Total[8] += (ytdBudget * switchpositivenegative);
                if (blockname.Length == 0) provisionalTax[9] = (ytdBudget * switchpositivenegative);
                //column 12
                datarow[11] = NegativeCurrencyFormat((thisPeriodLastYear.GetAmount(currency) * switchpositivenegative));
                allsummary[block, 10] += (thisPeriodLastYear.GetAmount(currency) * switchpositivenegative); 
                Column_Total[10] += (thisPeriodLastYear.GetAmount(currency) * switchpositivenegative);
                if (blockname.Length == 0) provisionalTax[11] = (thisPeriodLastYear.GetAmount(currency) * switchpositivenegative);

                if (!is_Surcharge) table.Rows.Add(datarow);

            }
            if (flagInsert && Implement_Changes)
            {
                //Add empty row
                DataRow emprow2 = table.NewRow();

                /*
                DataRow datarow2 = table.NewRow();
                datarow2[0] = "STEEL SURCHARGE";
                hasProductionSales = true;

                for (int i = 0; i <= 11; i++)
                {
                    datarow2[i + 1] = NegativeCurrencyFormat(Surcharge_Row[i]);
                }

                table.Rows.Add(datarow2);*/ // Depriciated and moved above

                DataRow datarow2 = table.NewRow();
                datarow2[0] = "TOTAL OTHER SALES";
                hasProductionSales = true;


                for (int i = 0; i <= 11; i++)
                {
                    datarow2[i + 1] = NegativeCurrencyFormat(Column_Total[i] + Surcharge_Row[i]);
                    Column_Total[i] = 0; // reset
                }

                table.Rows.Add(datarow2);
                total_line_row_number[2] = table.Rows.Count - 1;
                table.Rows.Add(emprow2);

            }


                /*
            else if (!hasOtherSales && group.name.Contains("INTERDIVISION SALES STEEL SURCHARGE")  && Implement_Changes)
            {
                sheet.Cells[row, 1] = "TOTAL STEEL SURCHARGE";
                for (int i = 2; i <= 13; i++)
                {
                    string colCode = Convert.ToChar((Convert.ToInt32('A') + i - 1)).ToString();
                    sheet.Cells[row, i].Formula = "=SUM(" + colCode + (SteelSurchargeRow + 1).ToString() + ":" + colCode + (row - 1).ToString() + ")";
                }
                row++;
                row++;
                sheet.Cells[row, 1] = "OTHER SALES";
                sheet.Cells.get_Range("A" + row.ToString()).Interior.ColorIndex = 40;
                row++;
                hasOtherSales = true;
            }
            else if (!hasTotalSales && group.name.Contains("SCRAP SALES") && Implement_Changes)
            {
                sheet.Cells[row, 1] = "TOTAL OTHER SALES";
                for (int i = 2; i <= 13; i++)
                {
                    string colCode = Convert.ToChar((Convert.ToInt32('A') + i - 1)).ToString();
                    sheet.Cells[row, i].Formula = "=SUM(" + colCode + (OtherSalesRow + 1).ToString() + ":" + colCode + (row - 1).ToString() + ")";
                }
                row++;
                hasTotalSales = true;
                //row++;
            }*/

        }

        if (blockname.Length > 0)
        {
            DataRow total = table.NewRow();
            total[0] = "TOTAL:";
            total[1] = NegativeCurrencyFormat(allsummary[block, 0]);
            total[3] = NegativeCurrencyFormat(allsummary[block, 2]);
            total[5] = NegativeCurrencyFormat(allsummary[block, 4]);
            total[7] = NegativeCurrencyFormat(allsummary[block, 6]);
            total[9] = NegativeCurrencyFormat(allsummary[block, 8]);
            total[11] = NegativeCurrencyFormat(allsummary[block, 10]);
            table.Rows.Add(total);
        }

        blockposition[block, 1] = table.Rows.Count - (blockname == "" ? 0 : 1);

        string line_no = "";
        try
        {
            //column 3,5,7,9,11
            for (int i = blockposition[block, 0]; i < blockposition[block, 1]; i++)
            {
                
                line_no = i.ToString();
                if (isZero(allsummary[0, 0]) == false && (table.Rows[i][1].ToString().Length > 0))
                {
                    tempdata = table.Rows[i][1].ToString();
                    tempp = double.Parse(tempdata, styles);
                    table.Rows[i][2] = (tempp / allsummary[0, 0]).ToString("P2");
                    allsummary[block, 1] += tempp / allsummary[0, 0];
                    Column_Total[1] += tempp / allsummary[0, 0];
                    if (blockname.Length == 0) provisionalTax[2] = tempp / allsummary[0, 0];
                }
                else if (table.Rows[i][1].ToString().Length > 0)
                {
                    table.Rows[i][2] = "0.00%";
                    allsummary[block, 1] = 0.00;
                    Column_Total[1] = 0.00;
                }

                if (isZero(allsummary[0, 2]) == false && (table.Rows[i][3].ToString().Length > 0))
                {
                    tempdata = table.Rows[i][3].ToString();
                    tempp = double.Parse(tempdata, styles);
                    table.Rows[i][4] = (tempp / allsummary[0, 2]).ToString("P2");
                    allsummary[block, 3] += tempp / allsummary[0, 2];
                    Column_Total[3] += tempp / allsummary[0, 2];
                    if (blockname.Length == 0) provisionalTax[4] = tempp / allsummary[0, 2];
                }
                else if (table.Rows[i][3].ToString().Length > 0)
                {
                    table.Rows[i][4] = "0.00%";
                    allsummary[block, 3] = 0.00;
                    Column_Total[3] = 0.00;
                }

                if (isZero(allsummary[0, 4]) == false && (table.Rows[i][5].ToString().Length > 0))
                {
                    tempdata = table.Rows[i][5].ToString();
                    tempp = double.Parse(tempdata, styles);
                    table.Rows[i][6] = (tempp / allsummary[0, 4]).ToString("P2");
                    allsummary[block, 5] += tempp / allsummary[0, 4];
                    Column_Total[5] += tempp / allsummary[0, 4];
                    if (blockname.Length == 0) provisionalTax[6] = tempp / allsummary[0, 4];
                }
                else if (table.Rows[i][5].ToString().Length > 0)
                {
                    table.Rows[i][6] = "0.00%";
                    allsummary[block, 5] = 0.00;
                    Column_Total[5] = 0.00;
                }

                if (isZero(allsummary[0, 6]) == false && (table.Rows[i][7].ToString().Length > 0))
                {
                    tempdata = table.Rows[i][7].ToString();
                    tempp = double.Parse(tempdata, styles);
                    table.Rows[i][8] = (tempp / allsummary[0, 6]).ToString("P2");
                    allsummary[block, 7] += tempp / allsummary[0, 6];
                    Column_Total[7] += tempp / allsummary[0, 6];
                    if (blockname.Length == 0) provisionalTax[8] = tempp / allsummary[0, 6];
                }
                else if (table.Rows[i][7].ToString().Length > 0)
                {
                    table.Rows[i][8] = "0.00%";
                    allsummary[block, 7] = 0.00;
                    Column_Total[7] = 0.00;
                }

                if (isZero(allsummary[0, 8]) == false && (table.Rows[i][9].ToString().Length > 0))
                {
                    tempdata = table.Rows[i][9].ToString();
                    tempp = double.Parse(tempdata, styles);
                    table.Rows[i][10] = (tempp / allsummary[0, 8]).ToString("P2");
                    allsummary[block, 9] += tempp / allsummary[0, 8];
                    Column_Total[9] += tempp / allsummary[0, 8];
                    if (blockname.Length == 0) provisionalTax[10] = tempp / allsummary[0, 8];
                }
                else if (table.Rows[i][9].ToString().Length > 0)
                {
                    table.Rows[i][10] = "0.00%";
                    allsummary[block, 9] = 0.00;
                    Column_Total[9] = 0.00;
                }

                if (isZero(allsummary[0, 10]) == false && (table.Rows[i][11].ToString().Length > 0))
                {
                    tempdata = table.Rows[i][11].ToString();
                    tempp = double.Parse(tempdata, styles);
                    table.Rows[i][12] = (tempp / allsummary[0, 10]).ToString("P2");
                    allsummary[block, 11] += tempp / allsummary[0, 10];
                    Column_Total[11] += tempp / allsummary[0, 10];
                    if (blockname.Length == 0) provisionalTax[12] = tempp / allsummary[0, 10];
                }
                else if (table.Rows[i][11].ToString().Length > 0)
                {
                    table.Rows[i][12] = "0.00%";
                    allsummary[block, 11] = 0.00;
                    Column_Total[11] = 0.00;
                }
            }
        }
        catch
        {
            Console.WriteLine("error at: " + line_no);
        }

        if (blockname.Length > 0)
        {
            table.Rows[blockposition[block, 1]][2] = (allsummary[block, 1] > 1 ? allsummary[block, 1] - 1 : allsummary[block, 1]).ToString("P2");
            table.Rows[blockposition[block, 1]][4] = (allsummary[block, 3] > 1 ? allsummary[block, 3] - 1 : allsummary[block, 3]).ToString("P2");
            table.Rows[blockposition[block, 1]][6] =( allsummary[block, 5] > 1 ? allsummary[block, 5] - 1 : allsummary[block, 5]).ToString("P2");
            table.Rows[blockposition[block, 1]][8] = (allsummary[block, 7] > 1 ? allsummary[block, 7] - 1 : allsummary[block, 7]).ToString("P2");
            table.Rows[blockposition[block, 1]][10] = (allsummary[block, 9] > 1 ? allsummary[block, 9] - 1 : allsummary[block, 9]).ToString("P2");
            table.Rows[blockposition[block, 1]][12] = (allsummary[block, 11] > 1 ? allsummary[block, 11] - 1 : allsummary[block, 11]).ToString("P2");
        }
        else
        {
            table.Rows[table.Rows.Count - 1][2] = provisionalTax[2].ToString("P2");
            table.Rows[table.Rows.Count - 1][4] = provisionalTax[4].ToString("P2");
            table.Rows[table.Rows.Count - 1][6] = provisionalTax[6].ToString("P2");
            table.Rows[table.Rows.Count - 1][8] = provisionalTax[8].ToString("P2");
            table.Rows[table.Rows.Count - 1][10] = provisionalTax[10].ToString("P2");
            table.Rows[table.Rows.Count - 1][12] = provisionalTax[12].ToString("P2");
        }
        
        //used to separate each block
        DataRow emprow = table.NewRow();
        table.Rows.Add(emprow);  
    }

    private void WriteTotalLineConsolidate(DataTable table, int totalID)
    {
        if (totalID == 0)
        {
            DataRow blockrow = table.NewRow();
            blockrow["Name"] = "TOTAL COST OF GOODS:";
            table.Rows.Add(blockrow);

            totalposition[0] = table.Rows.Count - 1;

            table.Rows[totalposition[0]][1] = NegativeCurrencyFormat(allsummary[1, 0] + allsummary[2, 0] + allsummary[3, 0]);
            table.Rows[totalposition[0]][3] = NegativeCurrencyFormat(allsummary[1, 2] + allsummary[2, 2] + allsummary[3, 2]);
            table.Rows[totalposition[0]][5] = NegativeCurrencyFormat(allsummary[1, 4] + allsummary[2, 4] + allsummary[3, 4]);
            table.Rows[totalposition[0]][7] = NegativeCurrencyFormat(allsummary[1, 6] + allsummary[2, 6] + allsummary[3, 6]);
            table.Rows[totalposition[0]][9] = NegativeCurrencyFormat(allsummary[1, 8] + allsummary[2, 8] + allsummary[3, 8]);
            table.Rows[totalposition[0]][11] = NegativeCurrencyFormat(allsummary[1, 10] + allsummary[2, 10] + allsummary[3, 10]);
            
            table.Rows[totalposition[0]][2] = (allsummary[1, 1] + allsummary[2, 1] + allsummary[3, 1]).ToString("P2");
            table.Rows[totalposition[0]][4] = (allsummary[1, 3] + allsummary[2, 3] + allsummary[3, 3]).ToString("P2");
            table.Rows[totalposition[0]][6] = (allsummary[1, 5] + allsummary[2, 5] + allsummary[3, 5]).ToString("P2");
            table.Rows[totalposition[0]][8] = (allsummary[1, 7] + allsummary[2, 7] + allsummary[3, 7]).ToString("P2");
            table.Rows[totalposition[0]][10] = (allsummary[1, 9] + allsummary[2, 9] + allsummary[3, 9]).ToString("P2");            

            //used to separate each block
            DataRow emprow = table.NewRow();
            table.Rows.Add(emprow);
        }
        else if (totalID == 1)
        {
            DataRow blockrow = table.NewRow();
            blockrow["Name"] = "GROSS MARGIN:";
            table.Rows.Add(blockrow);

            totalposition[1] = table.Rows.Count - 1;

            table.Rows[totalposition[1]][1] = NegativeCurrencyFormat(allsummary[0, 0] - (allsummary[1, 0] + allsummary[2, 0] + allsummary[3, 0]));
            table.Rows[totalposition[1]][3] = NegativeCurrencyFormat(allsummary[0, 2] - (allsummary[1, 2] + allsummary[2, 2] + allsummary[3, 2]));
            table.Rows[totalposition[1]][5] = NegativeCurrencyFormat(allsummary[0, 4] - (allsummary[1, 4] + allsummary[2, 4] + allsummary[3, 4]));
            table.Rows[totalposition[1]][7] = NegativeCurrencyFormat(allsummary[0, 6] - (allsummary[1, 6] + allsummary[2, 6] + allsummary[3, 6]));
            table.Rows[totalposition[1]][9] = NegativeCurrencyFormat(allsummary[0, 8] - (allsummary[1, 8] + allsummary[2, 8] + allsummary[3, 8]));          
            table.Rows[totalposition[1]][11] = NegativeCurrencyFormat(allsummary[0, 10] - (allsummary[1, 10] + allsummary[2, 10] + allsummary[3, 10]));          

            double percent =  (allsummary[0, 1] - (allsummary[1, 1] + allsummary[2, 1] + allsummary[3, 1] + 1));
            table.Rows[totalposition[1]][2] = (percent).ToString("P2");
            table.Rows[totalposition[1]][4] = (allsummary[0, 3] - (allsummary[1, 3] + allsummary[2, 3] + allsummary[3, 3] + 1)).ToString("P2");
            table.Rows[totalposition[1]][6] = (allsummary[0, 5] - (allsummary[1, 5] + allsummary[2, 5] + allsummary[3, 5] + 1)).ToString("P2");
            table.Rows[totalposition[1]][8] = (allsummary[0, 7] - (allsummary[1, 7] + allsummary[2, 7] + allsummary[3, 7] + 1)).ToString("P2");
            table.Rows[totalposition[1]][10] = (allsummary[0, 9] - (allsummary[1, 9] + allsummary[2, 9] + allsummary[3, 9] + 1)).ToString("P2");           

            //used to separate each block
            DataRow emprow = table.NewRow();
            table.Rows.Add(emprow);
        }
        else if (totalID == 2)
        {
            DataRow blockrow = table.NewRow();
            blockrow["Name"] = "TOTAL EXPENSES:";
            table.Rows.Add(blockrow);

            totalposition[2] = table.Rows.Count - 1;

            table.Rows[totalposition[2]][1] = NegativeCurrencyFormat(allsummary[4, 0] + allsummary[5, 0] + allsummary[6, 0]);
            table.Rows[totalposition[2]][3] = NegativeCurrencyFormat(allsummary[4, 2] + allsummary[5, 2] + allsummary[6, 2]);
            table.Rows[totalposition[2]][5] = NegativeCurrencyFormat(allsummary[4, 4] + allsummary[5, 4] + allsummary[6, 4]);
            table.Rows[totalposition[2]][7] = NegativeCurrencyFormat(allsummary[4, 6] + allsummary[5, 6] + allsummary[6, 6]);
            table.Rows[totalposition[2]][9] = NegativeCurrencyFormat(allsummary[4, 8] + allsummary[5, 8] + allsummary[6, 8]);
            table.Rows[totalposition[2]][11] = NegativeCurrencyFormat(allsummary[4, 10] + allsummary[5, 10] + allsummary[6, 10]);

            table.Rows[totalposition[2]][2] = (allsummary[4, 1] + allsummary[5, 1] + allsummary[6, 1]).ToString("P2");
            table.Rows[totalposition[2]][4] = (allsummary[4, 3] + allsummary[5, 3] + allsummary[6, 3]).ToString("P2");
            table.Rows[totalposition[2]][6] = (allsummary[4, 5] + allsummary[5, 5] + allsummary[6, 5]).ToString("P2");
            table.Rows[totalposition[2]][8] = (allsummary[4, 7] + allsummary[5, 7] + allsummary[6, 7]).ToString("P2");
            table.Rows[totalposition[2]][10] = (allsummary[4, 9] + allsummary[5, 9] + allsummary[6, 9]).ToString("P2");

            //used to separate each block
            DataRow emprow = table.NewRow();
            table.Rows.Add(emprow);
        }
        else if (totalID == 3)
        {
            DataRow blockrow = table.NewRow();
            blockrow["Name"] = "INCOME BEFORE PROVISION TAX:";
            table.Rows.Add(blockrow);

            totalposition[3] = table.Rows.Count - 1;

            table.Rows[totalposition[3]][1] = NegativeCurrencyFormat(allsummary[0, 0] - (allsummary[1, 0] + allsummary[2, 0] + allsummary[3, 0] + allsummary[4, 0] + allsummary[5, 0] + allsummary[6, 0]));
            table.Rows[totalposition[3]][3] = NegativeCurrencyFormat(allsummary[0, 2] - (allsummary[1, 2] + allsummary[2, 2] + allsummary[3, 2] + allsummary[4, 2] + allsummary[5, 2] + allsummary[6, 2]));
            table.Rows[totalposition[3]][5] = NegativeCurrencyFormat(allsummary[0, 4] - (allsummary[1, 4] + allsummary[2, 4] + allsummary[3, 4] + allsummary[4, 4] + allsummary[5, 4] + allsummary[6, 4]));
            table.Rows[totalposition[3]][7] = NegativeCurrencyFormat(allsummary[0, 6] - (allsummary[1, 6] + allsummary[2, 6] + allsummary[3, 6] + allsummary[4, 6] + allsummary[5, 6] + allsummary[6, 6]));
            table.Rows[totalposition[3]][9] = NegativeCurrencyFormat(allsummary[0, 8] - (allsummary[1, 8] + allsummary[2, 8] + allsummary[3, 8] + allsummary[4, 8] + allsummary[5, 8] + allsummary[6, 8]));
            table.Rows[totalposition[3]][11] = NegativeCurrencyFormat(allsummary[0, 10] - (allsummary[1, 10] + allsummary[2, 10] + allsummary[3, 10] + allsummary[4, 10] + allsummary[5, 10] + allsummary[6, 10]));

            table.Rows[totalposition[3]][2] = ((allsummary[0, 1] - (allsummary[1, 1] + allsummary[2, 1] + allsummary[3, 1] + allsummary[4, 1] + allsummary[5, 1] + allsummary[6, 1] + 1))).ToString("P2");
            table.Rows[totalposition[3]][4] = ((allsummary[0, 3] - (allsummary[1, 3] + allsummary[2, 3] + allsummary[3, 3] + allsummary[4, 3] + allsummary[5, 3] + allsummary[6, 3] + 1))).ToString("P2");
            table.Rows[totalposition[3]][6] = ((allsummary[0, 5] - (allsummary[1, 5] + allsummary[2, 5] + allsummary[3, 5] + allsummary[4, 5] + allsummary[5, 5] + allsummary[6, 5] + 1))).ToString("P2");
            table.Rows[totalposition[3]][8] = ((allsummary[0, 7] - (allsummary[1, 7] + allsummary[2, 7] + allsummary[3, 7] + allsummary[4, 7] + allsummary[5, 7] + allsummary[6, 7] + 1))).ToString("P2");
            table.Rows[totalposition[3]][10] = ((allsummary[0, 9] -( allsummary[1, 9] + allsummary[2, 9] + allsummary[3, 9] + allsummary[4, 9] + allsummary[5, 9] + allsummary[6, 9] + 1))).ToString("P2");

            //used to separate each block
            //DataRow emprow = table.NewRow();
            //table.Rows.Add(emprow);
        }
        else if (totalID == 4)
        {
            DataRow blockrow = table.NewRow();
            blockrow["Name"] = "NET INCOME";
            table.Rows.Add(blockrow);

            totalposition[4] = table.Rows.Count - 1;

            table.Rows[totalposition[4]][1] = NegativeCurrencyFormat(-provisionalTax[1] + allsummary[0, 0] - (allsummary[1, 0] + allsummary[2, 0] + allsummary[3, 0] + allsummary[4, 0] + allsummary[5, 0] + allsummary[6, 0]));
            table.Rows[totalposition[4]][3] = NegativeCurrencyFormat(-provisionalTax[3] + allsummary[0, 2] - (allsummary[1, 2] + allsummary[2, 2] + allsummary[3, 2] + allsummary[4, 2] + allsummary[5, 2] + allsummary[6, 2]));
            table.Rows[totalposition[4]][5] = NegativeCurrencyFormat(-provisionalTax[5] + allsummary[0, 4] - (allsummary[1, 4] + allsummary[2, 4] + allsummary[3, 4] + allsummary[4, 4] + allsummary[5, 4] + allsummary[6, 4]));
            table.Rows[totalposition[4]][7] = NegativeCurrencyFormat(-provisionalTax[7] + allsummary[0, 6] - (allsummary[1, 6] + allsummary[2, 6] + allsummary[3, 6] + allsummary[4, 6] + allsummary[5, 6] + allsummary[6, 6]));
            table.Rows[totalposition[4]][9] = NegativeCurrencyFormat(-provisionalTax[9] + allsummary[0, 8] - (allsummary[1, 8] + allsummary[2, 8] + allsummary[3, 8] + allsummary[4, 8] + allsummary[5, 8] + allsummary[6, 8]));
            table.Rows[totalposition[4]][11] = NegativeCurrencyFormat(-provisionalTax[11] + allsummary[0, 10] - (allsummary[1, 10] + allsummary[2, 10] + allsummary[3, 10] + allsummary[4, 10] + allsummary[5, 10] + allsummary[6, 10]));

            table.Rows[totalposition[4]][2] = (-(provisionalTax[1] / allsummary[0, 0]) + allsummary[0, 1] - (allsummary[1, 1] + allsummary[2, 1] + allsummary[3, 1] + allsummary[4, 1] + allsummary[5, 1] + allsummary[6, 1] + 1)).ToString("P2");
            table.Rows[totalposition[4]][4] = (-(provisionalTax[3] / allsummary[0, 2]) + allsummary[0, 3] - (allsummary[1, 3] + allsummary[2, 3] + allsummary[3, 3] + allsummary[4, 3] + allsummary[5, 3] + allsummary[6, 3] + 1)).ToString("P2");
            table.Rows[totalposition[4]][6] = (-(provisionalTax[5] / allsummary[0, 4]) + allsummary[0, 5] - (allsummary[1, 5] + allsummary[2, 5] + allsummary[3, 5] + allsummary[4, 5] + allsummary[5, 5] + allsummary[6, 5] + 1)).ToString("P2");
            table.Rows[totalposition[4]][8] = (-(provisionalTax[7] / allsummary[0, 6]) + allsummary[0, 7] - (allsummary[1, 7] + allsummary[2, 7] + allsummary[3, 7] + allsummary[4, 7] + allsummary[5, 7] + allsummary[6, 7] + 1)).ToString("P2");
            table.Rows[totalposition[4]][10] = (-(provisionalTax[9] / allsummary[0, 8]) + allsummary[0, 9] - (allsummary[1, 9] + allsummary[2, 9] + allsummary[3, 9] + allsummary[4, 9] + allsummary[5, 9] + allsummary[6, 9] + 1)).ToString("P2");

            //used to separate each block
            DataRow emprow = table.NewRow();
            table.Rows.Add(emprow);
        }
    }

    private void WriteTotalLine(DataTable table, int totalID)
    {
        if (totalID == 0)
        {
            DataRow blockrow = table.NewRow();
            blockrow["Name"] = "TOTAL COST OF GOODS:";
            table.Rows.Add(blockrow);  

            totalposition[0] = table.Rows.Count - 1;

            table.Rows[totalposition[0]][1] = NegativeCurrencyFormat(allsummary[1, 0] + allsummary[2, 0] + allsummary[3, 0]);
            table.Rows[totalposition[0]][3] = NegativeCurrencyFormat(allsummary[1, 2] + allsummary[2, 2] + allsummary[3, 2]);
            table.Rows[totalposition[0]][5] = NegativeCurrencyFormat(allsummary[1, 4] + allsummary[2, 4] + allsummary[3, 4]);
            table.Rows[totalposition[0]][7] = NegativeCurrencyFormat(allsummary[1, 6] + allsummary[2, 6] + allsummary[3, 6]);
            table.Rows[totalposition[0]][9] = NegativeCurrencyFormat(allsummary[1, 8] + allsummary[2, 8] + allsummary[3, 8]);
            table.Rows[totalposition[0]][11] = NegativeCurrencyFormat(allsummary[1, 10] + allsummary[2, 10] + allsummary[3, 10]);

            table.Rows[totalposition[0]][2] = (allsummary[1, 1] + allsummary[2, 1] + allsummary[3, 1]).ToString("P2");
            table.Rows[totalposition[0]][4] = (allsummary[1, 3] + allsummary[2, 3] + allsummary[3, 3]).ToString("P2");
            table.Rows[totalposition[0]][6] = (allsummary[1, 5] + allsummary[2, 5] + allsummary[3, 5]).ToString("P2");
            table.Rows[totalposition[0]][8] = (allsummary[1, 7] + allsummary[2, 7] + allsummary[3, 7]).ToString("P2");
            table.Rows[totalposition[0]][10] = (allsummary[1, 9] + allsummary[2, 9] + allsummary[3, 9]).ToString("P2");
            table.Rows[totalposition[0]][12] = (allsummary[1, 11] + allsummary[2, 11] + allsummary[3, 11]).ToString("P2");
           
            //used to separate each block
            DataRow emprow = table.NewRow();
            table.Rows.Add(emprow);              
        }
        else if (totalID == 1)
        {
            DataRow blockrow = table.NewRow();
            blockrow["Name"] = "GROSS MARGIN:";
            table.Rows.Add(blockrow);

            totalposition[1] = table.Rows.Count - 1;

            table.Rows[totalposition[1]][1] = NegativeCurrencyFormat(allsummary[0, 0] + allsummary[1, 0] + allsummary[2, 0] + allsummary[3, 0]);
            table.Rows[totalposition[1]][3] = NegativeCurrencyFormat(allsummary[0, 2] + allsummary[1, 2] + allsummary[2, 2] + allsummary[3, 2]);
            table.Rows[totalposition[1]][5] = NegativeCurrencyFormat(allsummary[0, 4] + allsummary[1, 4] + allsummary[2, 4] + allsummary[3, 4]);
            table.Rows[totalposition[1]][7] = NegativeCurrencyFormat(allsummary[0, 6] + allsummary[1, 6] + allsummary[2, 6] + allsummary[3, 6]);
            table.Rows[totalposition[1]][9] = NegativeCurrencyFormat(allsummary[0, 8] + allsummary[1, 8] + allsummary[2, 8] + allsummary[3, 8]);
            table.Rows[totalposition[1]][11] = NegativeCurrencyFormat(allsummary[0, 10] + allsummary[1, 10] + allsummary[2, 10] + allsummary[3, 10]);

            table.Rows[totalposition[1]][2] = (allsummary[0, 1] + allsummary[1, 1] + allsummary[2, 1] + allsummary[3, 1] - 1).ToString("P2");
            table.Rows[totalposition[1]][4] = (allsummary[0, 3] + allsummary[1, 3] + allsummary[2, 3] + allsummary[3, 3] - 1).ToString("P2");
            table.Rows[totalposition[1]][6] = (allsummary[0, 5] + allsummary[1, 5] + allsummary[2, 5] + allsummary[3, 5] - 1).ToString("P2");
            table.Rows[totalposition[1]][8] = (allsummary[0, 7] + allsummary[1, 7] + allsummary[2, 7] + allsummary[3, 7] - 1).ToString("P2");
            table.Rows[totalposition[1]][10] = (allsummary[0, 9] + allsummary[1, 9] + allsummary[2, 9] + allsummary[3, 9] - 1).ToString("P2");
            table.Rows[totalposition[1]][12] = (allsummary[0, 11] + allsummary[1, 11] + allsummary[2, 11] + allsummary[3, 11] - 1).ToString("P2");

            //used to separate each block
            DataRow emprow = table.NewRow();
            table.Rows.Add(emprow);   
        }
        else if (totalID == 2)
        {
            DataRow blockrow = table.NewRow();
            blockrow["Name"] = "TOTAL EXPENSES:";
            table.Rows.Add(blockrow);

            totalposition[2] = table.Rows.Count - 1;

            table.Rows[totalposition[2]][1] = NegativeCurrencyFormat(allsummary[4, 0] + allsummary[5, 0] + allsummary[6, 0]);
            table.Rows[totalposition[2]][3] = NegativeCurrencyFormat(allsummary[4, 2] + allsummary[5, 2] + allsummary[6, 2]);
            table.Rows[totalposition[2]][5] = NegativeCurrencyFormat(allsummary[4, 4] + allsummary[5, 4] + allsummary[6, 4]);
            table.Rows[totalposition[2]][7] = NegativeCurrencyFormat(allsummary[4, 6] + allsummary[5, 6] + allsummary[6, 6]);
            table.Rows[totalposition[2]][9] = NegativeCurrencyFormat(allsummary[4, 8] + allsummary[5, 8] + allsummary[6, 8]);
            table.Rows[totalposition[2]][11] = NegativeCurrencyFormat(allsummary[4, 10] + allsummary[5, 10] + allsummary[6, 10]);

            table.Rows[totalposition[2]][2] = (allsummary[4, 1] + allsummary[5, 1] + allsummary[6, 1]).ToString("P2");
            table.Rows[totalposition[2]][4] = (allsummary[4, 3] + allsummary[5, 3] + allsummary[6, 3]).ToString("P2");
            table.Rows[totalposition[2]][6] = (allsummary[4, 5] + allsummary[5, 5] + allsummary[6, 5]).ToString("P2");
            table.Rows[totalposition[2]][8] = (allsummary[4, 7] + allsummary[5, 7] + allsummary[6, 7]).ToString("P2");
            table.Rows[totalposition[2]][10] = (allsummary[4, 9] + allsummary[5, 9] + allsummary[6, 9]).ToString("P2");
            table.Rows[totalposition[2]][12] = (allsummary[4, 11] + allsummary[5, 11] + allsummary[6, 11]).ToString("P2");

            //used to separate each block
            DataRow emprow = table.NewRow();
            table.Rows.Add(emprow);
        }
        else if (totalID == 3)
        {
            DataRow blockrow = table.NewRow();
            blockrow["Name"] = "INCOME BEFORE PROVISION TAX:";
            table.Rows.Add(blockrow);

            totalposition[3] = table.Rows.Count - 1;

            table.Rows[totalposition[3]][1] = NegativeCurrencyFormat(allsummary[0, 0] + allsummary[1, 0] + allsummary[2, 0] + allsummary[3, 0] + allsummary[4, 0] + allsummary[5, 0] + allsummary[6, 0]);
            table.Rows[totalposition[3]][3] = NegativeCurrencyFormat(allsummary[0, 2] + allsummary[1, 2] + allsummary[2, 2] + allsummary[3, 2] + allsummary[4, 2] + allsummary[5, 2] + allsummary[6, 2]);
            table.Rows[totalposition[3]][5] = NegativeCurrencyFormat(allsummary[0, 4] + allsummary[1, 4] + allsummary[2, 4] + allsummary[3, 4] + allsummary[4, 4] + allsummary[5, 4] + allsummary[6, 4]);
            table.Rows[totalposition[3]][7] = NegativeCurrencyFormat(allsummary[0, 6] + allsummary[1, 6] + allsummary[2, 6] + allsummary[3, 6] + allsummary[4, 6] + allsummary[5, 6] + allsummary[6, 6]);
            table.Rows[totalposition[3]][9] = NegativeCurrencyFormat(allsummary[0, 8] + allsummary[1, 8] + allsummary[2, 8] + allsummary[3, 8] + allsummary[4, 8] + allsummary[5, 8] + allsummary[6, 8]);
            table.Rows[totalposition[3]][11] = NegativeCurrencyFormat(allsummary[0, 10] + allsummary[1, 10] + allsummary[2, 10] + allsummary[3, 10] + allsummary[4, 10] + allsummary[5, 10] + allsummary[6, 10]);

            table.Rows[totalposition[3]][2] = (allsummary[0, 1] + allsummary[1, 1] + allsummary[2, 1] + allsummary[3, 1] + allsummary[4, 1] + allsummary[5, 1] + allsummary[6, 1] - 1).ToString("P2");
            table.Rows[totalposition[3]][4] = (allsummary[0, 3] + allsummary[1, 3] + allsummary[2, 3] + allsummary[3, 3] + allsummary[4, 3] + allsummary[5, 3] + allsummary[6, 3] - 1).ToString("P2");
            table.Rows[totalposition[3]][6] = (allsummary[0, 5] + allsummary[1, 5] + allsummary[2, 5] + allsummary[3, 5] + allsummary[4, 5] + allsummary[5, 5] + allsummary[6, 5] - 1).ToString("P2");
            table.Rows[totalposition[3]][8] = (allsummary[0, 7] + allsummary[1, 7] + allsummary[2, 7] + allsummary[3, 7] + allsummary[4, 7] + allsummary[5, 7] + allsummary[6, 7] - 1).ToString("P2");
            table.Rows[totalposition[3]][10] = (allsummary[0, 9] + allsummary[1, 9] + allsummary[2, 9] + allsummary[3, 9] + allsummary[4, 9] + allsummary[5, 9] + allsummary[6, 9] - 1).ToString("P2");
            table.Rows[totalposition[3]][12] = (allsummary[0, 11] + allsummary[1, 11] + allsummary[2, 11] + allsummary[3, 11] + allsummary[4, 11] + allsummary[5, 11] + allsummary[6, 11] - 1).ToString("P2");

            //used to separate each block
            DataRow emprow = table.NewRow();
            table.Rows.Add(emprow);
        }
        else if (totalID == 4)
        {
            DataRow blockrow = table.NewRow();
            blockrow["Name"] = "NET INCOME";
            table.Rows.Add(blockrow);

            totalposition[4] = table.Rows.Count - 1;

            table.Rows[totalposition[4]][1] = NegativeCurrencyFormat(provisionalTax[1] + allsummary[0, 0] + allsummary[1, 0] + allsummary[2, 0] + allsummary[3, 0] + allsummary[4, 0] + allsummary[5, 0] + allsummary[6, 0]);
            table.Rows[totalposition[4]][3] = NegativeCurrencyFormat(provisionalTax[3] + allsummary[0, 2] + allsummary[1, 2] + allsummary[2, 2] + allsummary[3, 2] + allsummary[4, 2] + allsummary[5, 2] + allsummary[6, 2]);
            table.Rows[totalposition[4]][5] = NegativeCurrencyFormat(provisionalTax[5] + allsummary[0, 4] + allsummary[1, 4] + allsummary[2, 4] + allsummary[3, 4] + allsummary[4, 4] + allsummary[5, 4] + allsummary[6, 4]);
            table.Rows[totalposition[4]][7] = NegativeCurrencyFormat(provisionalTax[7] + allsummary[0, 6] + allsummary[1, 6] + allsummary[2, 6] + allsummary[3, 6] + allsummary[4, 6] + allsummary[5, 6] + allsummary[6, 6]);
            table.Rows[totalposition[4]][9] = NegativeCurrencyFormat(provisionalTax[9] + allsummary[0, 8] + allsummary[1, 8] + allsummary[2, 8] + allsummary[3, 8] + allsummary[4, 8] + allsummary[5, 8] + allsummary[6, 8]);
            table.Rows[totalposition[4]][11] = NegativeCurrencyFormat(provisionalTax[11] + allsummary[0, 10] + allsummary[1, 10] + allsummary[2, 10] + allsummary[3, 10] + allsummary[4, 10] + allsummary[5, 10] + allsummary[6, 10]);

            table.Rows[totalposition[4]][2] = (provisionalTax[2] + allsummary[0, 1] + allsummary[1, 1] + allsummary[2, 1] + allsummary[3, 1] + allsummary[4, 1] + allsummary[5, 1] + allsummary[6, 1] - 1).ToString("P2");
            table.Rows[totalposition[4]][4] = (provisionalTax[4] + allsummary[0, 3] + allsummary[1, 3] + allsummary[2, 3] + allsummary[3, 3] + allsummary[4, 3] + allsummary[5, 3] + allsummary[6, 3] - 1).ToString("P2");
            table.Rows[totalposition[4]][6] = (provisionalTax[6] + allsummary[0, 5] + allsummary[1, 5] + allsummary[2, 5] + allsummary[3, 5] + allsummary[4, 5] + allsummary[5, 5] + allsummary[6, 5] - 1).ToString("P2");
            table.Rows[totalposition[4]][8] = (provisionalTax[8] + allsummary[0, 7] + allsummary[1, 7] + allsummary[2, 7] + allsummary[3, 7] + allsummary[4, 7] + allsummary[5, 7] + allsummary[6, 7] - 1).ToString("P2");
            table.Rows[totalposition[4]][10] = (provisionalTax[10] + allsummary[0, 9] + allsummary[1, 9] + allsummary[2, 9] + allsummary[3, 9] + allsummary[4, 9] + allsummary[5, 9] + allsummary[6, 9] - 1).ToString("P2");
            table.Rows[totalposition[4]][12] = (provisionalTax[12] + allsummary[0, 11] + allsummary[1, 11] + allsummary[2, 11] + allsummary[3, 11] + allsummary[4, 11] + allsummary[5, 11] + allsummary[6, 11] - 1).ToString("P2");

            //used to separate each block
            DataRow emprow = table.NewRow();
            table.Rows.Add(emprow);
        }
    }

    private void TableStyleAdjustment(DataTable table, GridView gridView)
    {
        //stlye type
        Style headstyle = new Style();
        headstyle.BackColor = Color.Yellow;
        headstyle.Font.Size = 12;

        //stlye type
        Style headstyle2 = new Style();
        headstyle2.BackColor = ColorTranslator.FromHtml("#FFCC99");
        headstyle2.Font.Bold = true;
        headstyle2.Font.Size = 9;

        Style totalstyle = new Style();
        totalstyle.BackColor = Color.LightGray;
        totalstyle.Font.Bold = true;
        totalstyle.Font.Size = 9;
        
        //adjust style
        foreach (GridViewRow row in gridView.Rows)
        {
            row.Cells[0].Wrap = false;
            row.Cells[0].HorizontalAlign = HorizontalAlign.Left;
            row.Cells[1].Wrap = false;
            row.Cells[1].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[2].Wrap = false;
            row.Cells[2].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[3].Wrap = false;
            row.Cells[3].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[4].Wrap = false;
            row.Cells[4].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[5].Wrap = false;
            row.Cells[5].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[6].Wrap = false;
            row.Cells[6].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[7].Wrap = false;
            row.Cells[7].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[8].Wrap = false;
            row.Cells[8].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[9].Wrap = false;
            row.Cells[9].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[10].Wrap = false;
            row.Cells[10].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[11].Wrap = false;
            row.Cells[11].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[12].Wrap = false;
            row.Cells[12].HorizontalAlign = HorizontalAlign.Right;
        }


        

        //block header Cell
        GridViewRow headrow = gridView.Rows[blockposition[0, 0] - 2];
        headrow.Cells[0].ApplyStyle(headstyle);
        headrow = gridView.Rows[blockposition[1, 0]];
        headrow.Cells[0].ApplyStyle(headstyle);
        headrow = gridView.Rows[blockposition[2, 0]];
        headrow.Cells[0].ApplyStyle(headstyle);
        headrow = gridView.Rows[blockposition[3, 0]];
        headrow.Cells[0].ApplyStyle(headstyle);
        headrow = gridView.Rows[blockposition[4, 0]];
        headrow.Cells[0].ApplyStyle(headstyle);
        headrow = gridView.Rows[blockposition[5, 0]];
        headrow.Cells[0].ApplyStyle(headstyle);
        headrow = gridView.Rows[blockposition[6, 0]];
        headrow.Cells[0].ApplyStyle(headstyle);

        //subcategory header cell
        headrow = gridView.Rows[line_subgroup[0]];
        headrow.Cells[0].ApplyStyle(headstyle2);
        //headrow = gridView.Rows[line_subgroup[1]];
        //headrow.Cells[0].ApplyStyle(headstyle2);
        headrow = gridView.Rows[line_subgroup[2]];
        headrow.Cells[0].ApplyStyle(headstyle2);


        //block total Row
        GridViewRow totalrow = gridView.Rows[blockposition[0, 1]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[blockposition[1, 1]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[blockposition[2, 1]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[blockposition[3, 1]];
        totalrow.ApplyStyle(totalstyle);

        totalrow = gridView.Rows[totalposition[0]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[totalposition[1]];
        totalrow.ApplyStyle(totalstyle);

        totalrow = gridView.Rows[blockposition[4, 1]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[blockposition[5, 1]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[blockposition[6, 1]];
        totalrow.ApplyStyle(totalstyle);

        totalrow = gridView.Rows[totalposition[2]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[totalposition[3]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[totalposition[4]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[total_line_row_number[0]];
        totalrow.ApplyStyle(totalstyle);
        //totalrow = gridView.Rows[total_line_row_number[1]];
        //totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[total_line_row_number[2]];
        totalrow.ApplyStyle(totalstyle);
    }

    private void TableStyleAdjustmentConsolidate(DataTable table, GridView gridView)
    {
        //stlye type
        Style headstyle = new Style();
        headstyle.BackColor = Color.Yellow;
        headstyle.Font.Size = 12;

        //stlye type
        Style headstyle2 = new Style();
        headstyle2.BackColor = ColorTranslator.FromHtml("#FFCC99");
        headstyle2.Font.Bold = true;
        headstyle2.Font.Size = 9;

        Style totalstyle = new Style();
        totalstyle.BackColor = Color.LightGray;
        totalstyle.Font.Bold = true;
        totalstyle.Font.Size = 9;

        Style headconsolide = new Style();
        headconsolide.Font.Bold = true;
        headconsolide.Font.Size = 9;

        Style SupplyInfo = new Style();
        SupplyInfo.BackColor = Color.LightBlue;
        SupplyInfo.Font.Bold = true;
        SupplyInfo.Font.Size = 9;

        //adjust style
        foreach (GridViewRow row in gridView.Rows)
        {
            row.Cells[0].Wrap = false;
            row.Cells[0].HorizontalAlign = HorizontalAlign.Left;
            row.Cells[1].Wrap = false;
            row.Cells[1].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[2].Wrap = false;
            row.Cells[2].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[3].Wrap = false;
            row.Cells[3].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[4].Wrap = false;
            row.Cells[4].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[5].Wrap = false;
            row.Cells[5].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[6].Wrap = false;
            row.Cells[6].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[7].Wrap = false;
            row.Cells[7].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[8].Wrap = false;
            row.Cells[8].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[9].Wrap = false;
            row.Cells[9].HorizontalAlign = HorizontalAlign.Right;
            row.Cells[10].Wrap = false;
            row.Cells[10].HorizontalAlign = HorizontalAlign.Right;
        }
        //block header Cell
        GridViewRow headrow = gridView.Rows[blockposition[0, 0]-2];
        headrow.Cells[0].ApplyStyle(headstyle);
        headrow = gridView.Rows[blockposition[1, 0]];
        headrow.Cells[0].ApplyStyle(headstyle);
        headrow = gridView.Rows[blockposition[2, 0]];
        headrow.Cells[0].ApplyStyle(headstyle);
        headrow = gridView.Rows[blockposition[3, 0]];
        headrow.Cells[0].ApplyStyle(headstyle);
        headrow = gridView.Rows[blockposition[4, 0]];
        headrow.Cells[0].ApplyStyle(headstyle);
        headrow = gridView.Rows[blockposition[5, 0]];
        headrow.Cells[0].ApplyStyle(headstyle);
        headrow = gridView.Rows[blockposition[6, 0]];
        headrow.Cells[0].ApplyStyle(headstyle);

        // Subcategories
        headrow = gridView.Rows[line_subgroup[0]];
        headrow.Cells[0].ApplyStyle(headstyle2);
        //headrow = gridView.Rows[line_subgroup[1]];
        //headrow.Cells[0].ApplyStyle(headstyle2);
        headrow = gridView.Rows[line_subgroup[2]];
        headrow.Cells[0].ApplyStyle(headstyle2);

        //block total Row
        GridViewRow totalrow = gridView.Rows[blockposition[0, 1]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[blockposition[1, 1]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[blockposition[2, 1]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[blockposition[3, 1]];
        totalrow.ApplyStyle(totalstyle);

        totalrow = gridView.Rows[totalposition[0]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[totalposition[1]];
        totalrow.ApplyStyle(totalstyle);

        //sales subtotals
        totalrow = gridView.Rows[total_line_row_number[0]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[total_line_row_number[1]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[total_line_row_number[2]];
        totalrow.ApplyStyle(totalstyle);

        totalrow = gridView.Rows[blockposition[4, 1]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[blockposition[5, 1]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[blockposition[6, 1]];
        totalrow.ApplyStyle(totalstyle);

        totalrow = gridView.Rows[totalposition[2]];
        totalrow.ApplyStyle(totalstyle);
        totalrow = gridView.Rows[totalposition[3]];
        totalrow.ApplyStyle(totalstyle);

        totalrow = gridView.Rows[totalposition[4]];
        totalrow.ApplyStyle(totalstyle);

        totalrow = gridView.Rows[totalposition[5]];
        totalrow.Cells[0].ApplyStyle(SupplyInfo);

        totalrow = gridView.Rows[totalposition[6]];
        totalrow.ApplyStyle(totalstyle);

        totalrow = gridView.Rows[totalposition[7]];
        totalrow.ApplyStyle(totalstyle);

        totalrow = gridView.Rows[totalposition[8]];
        totalrow.ApplyStyle(totalstyle);

        totalrow = gridView.Rows[totalposition[9]];
        totalrow.ApplyStyle(totalstyle);

        totalrow = gridView.Rows[totalposition[10]];
        totalrow.ApplyStyle(totalstyle);

        totalrow = gridView.Rows[totalposition[11]];
        totalrow.ApplyStyle(totalstyle);

        //other info
        //GridViewRow headcons = gridView.Rows[blockposition[7, 0]];
        //headcons.Cells[0].ApplyStyle(headconsolide);
        //headcons = gridView.Rows[blockposition[7, 1]];
        //headcons.ApplyStyle(totalstyle);
    }
}