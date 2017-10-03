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

public partial class Trial_Balance : System.Web.UI.Page
{
    public class Account
    {
        public int glNum1 = 0;
        public int glNum2 = 0;
        public string title = string.Empty;
        public double balanceAmount01 = 0;
        public double balanceAmount03 = 0;
        public double balanceAmount05 = 0;
        public double balanceAmount04 = 0;
        public double balanceAmount41 = 0;
        public double balanceAmount48 = 0;
        public double balanceAmount49 = 0;
    };

    Dictionary<int, Account> accountMap = new Dictionary<int, Account>();

    private DataTable dataTable = new DataTable();

    private int plantID = 0;

    protected void Page_Load(object sender, EventArgs e)
    {
        // process user privilege
        if (true == (bool)Session["_hasTrialBalanceTexas"])
        {
            SelectTexas.Visible = true;
        }
        if (true == (bool)Session["_hasTrialBalanceColombia"])
        {
            SelectColombia.Visible = true;
        }
        if (true == (bool)Session["_hasTrialBalanceMichigan"])
        {
            SelectMichigan.Visible = true;
        }
        if (true == (bool)Session["_hasTrialBalanceMarkham"])
        {
            SelectMarkham.Visible = true;
        }
        // load all fiscal years
        var connection = new OdbcConnection(ConfigurationManager.ConnectionStrings["cms1"].ConnectionString);
        connection.Open();
        var command = new OdbcCommand();
        command.Connection = connection;
        var query = "select distinct aj4ccyy from cmsdat.glmt order by aj4ccyy desc";
        command.CommandText = query;
        var reader = command.ExecuteReader();
        while (reader.Read())
        {
            YearList.Items.Add(reader[0].ToString().Trim());
        }
        reader.Close();
        connection.Close();
    }

    private void BuildDataTable()
    {
        // create grid view data table
        dataTable.Columns.Add("Account #");
        dataTable.Columns.Add("Account Name");
        // create period columns
        ExcoCalendar calendar = new ExcoCalendar(Convert.ToInt32(YearList.SelectedValue) - 2000, PeriodList.SelectedIndex + 1, true, 1);
        if (SelectColombia.Checked)
        {
            calendar = new ExcoCalendar(Convert.ToInt32(YearList.SelectedValue) - 2000, PeriodList.SelectedIndex + 1, false, 1);
            dataTable.Columns.Add("04 Period " + calendar.GetCalendarYear().ToString("D2") + "-" + calendar.GetCalendarMonth().ToString("D2"));
            dataTable.Columns.Add("41 Period " + calendar.GetFiscalYear().ToString("D2") + "-" + calendar.GetFiscalMonth().ToString("D2"));
            dataTable.Columns.Add("48 Period " + calendar.GetCalendarYear().ToString("D2") + "-" + calendar.GetCalendarMonth().ToString("D2"));
            dataTable.Columns.Add("49 Period " + calendar.GetFiscalYear().ToString("D2") + "-" + calendar.GetFiscalMonth().ToString("D2"));
            dataTable.Columns.Add("Consolidated");
            plantID = 4;
        }
        else if (SelectMarkham.Checked)
        {
            dataTable.Columns.Add("01 Period " + calendar.GetFiscalYear().ToString("D2") + "-" + calendar.GetFiscalMonth().ToString("D2"));
            plantID = 1;
        }
        else if (SelectMichigan.Checked)
        {
            dataTable.Columns.Add("03 Period " + calendar.GetFiscalYear().ToString("D2") + "-" + calendar.GetFiscalMonth().ToString("D2"));
            plantID = 3;
        }
        else if (SelectTexas.Checked)
        {
            dataTable.Columns.Add("05 Period " + calendar.GetFiscalYear().ToString("D2") + "-" + calendar.GetFiscalMonth().ToString("D2"));
            plantID = 5;
        }
    }

    private void FillDataTable()
    {
        ExcoODBC database = ExcoODBC.Instance;
        database.Open(Database.CMSDAT);
        string query = string.Empty;
        OdbcDataReader reader;
        ExcoCalendar calendar = new ExcoCalendar(Convert.ToInt32(YearList.SelectedValue) - 2000, Convert.ToInt32(PeriodList.SelectedValue), true, plantID);
        ExcoCalendar tempCalendar = new ExcoCalendar(Convert.ToInt32(YearList.SelectedValue) - 2000, Convert.ToInt32(PeriodList.SelectedValue), false, 1);
        if (SelectColombia.Checked)
        {
            // plant 04
            query = "select a.aj4gl#1 as no1, a.aj4gl#2 as no2, b.aztitl as title, b.azatyp as account_type, (a.aj4tt" + calendar.GetFiscalMonth().ToString("D2") + "+aj4ob" + calendar.GetFiscalMonth().ToString("D2") + ") as balance from cmsdat.glmt as a, cmsdat.mast as b where a.aj4comp=4 and a.aj4ccyy=" + (calendar.GetFiscalYear() + 2000).ToString("D2") + " and a.aj4gl#1=b.azgl#1 and a.aj4gl#2=b.azgl#2 and b.azcomp=4";
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                Account account = new Account();
                account.glNum1 = Convert.ToInt32(reader["no1"]) % 100;
                account.glNum2 = Convert.ToInt32(reader["no2"]);
                account.title = reader["title"].ToString().Trim();
                account.balanceAmount04 = Convert.ToDouble(reader["balance"]);
                int key = account.glNum1 * 1000000 + account.glNum2;
                if (99999999 == key)
                {
                    continue;
                }
                if (accountMap.ContainsKey(key))
                {
                    accountMap[key].balanceAmount04 = account.balanceAmount04;
                }
                else
                {
                    accountMap.Add(key, account);
                }
            }
            reader.Close();
            // plant 48
            query = "select a.aj4gl#1 as no1, a.aj4gl#2 as no2, b.aztitl as title, b.azatyp as account_type, (a.aj4tt" + calendar.GetFiscalMonth().ToString("D2") + "+aj4ob" + calendar.GetFiscalMonth().ToString("D2") + ") as balance from cmsdat.glmt as a, cmsdat.mast as b where a.aj4comp=48 and a.aj4ccyy=" + (calendar.GetFiscalYear() + 2000).ToString("D2") + " and a.aj4gl#1=b.azgl#1 and a.aj4gl#2=b.azgl#2 and b.azcomp=48";
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                Account account = new Account();
                account.glNum1 = Convert.ToInt32(reader["no1"]) % 100;
                account.glNum2 = Convert.ToInt32(reader["no2"]);
                account.title = reader["title"].ToString().Trim();
                account.balanceAmount48 = Convert.ToDouble(reader["balance"]);
                int key = account.glNum1 * 1000000 + account.glNum2;
                if (99999999 == key)
                {
                    continue;
                }
                if (accountMap.ContainsKey(key))
                {
                    accountMap[key].balanceAmount48 = account.balanceAmount48;
                }
                else
                {
                    accountMap.Add(key, account);
                }
            }
            reader.Close();
            // plant 41
            query = "select a.aj4gl#1 as no1, a.aj4gl#2 as no2, b.aztitl as title, b.azatyp as account_type, (a.aj4tt" + tempCalendar.GetFiscalMonth().ToString("D2") + "+aj4ob" + tempCalendar.GetFiscalMonth().ToString("D2") + ") as balance from cmsdat.glmt as a, cmsdat.mast as b where a.aj4comp=41 and a.aj4ccyy=" + (tempCalendar.GetFiscalYear() + 2000).ToString("D2") + " and a.aj4gl#1=b.azgl#1 and a.aj4gl#2=b.azgl#2 and b.azcomp=41";
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                Account account = new Account();
                account.glNum1 = Convert.ToInt32(reader["no1"]) % 100;
                account.glNum2 = Convert.ToInt32(reader["no2"]);
                account.title = reader["title"].ToString().Trim();
                account.balanceAmount41 = Convert.ToDouble(reader["balance"]);
                int key = account.glNum1 * 1000000 + account.glNum2;
                if (99999999 == key)
                {
                    continue;
                }
                if (accountMap.ContainsKey(key))
                {
                    accountMap[key].balanceAmount41 = account.balanceAmount41;
                }
                else
                {
                    accountMap.Add(key, account);
                }
            }
            reader.Close();
            // plant 49
            query = "select a.aj4gl#1 as no1, a.aj4gl#2 as no2, b.aztitl as title, b.azatyp as account_type, (a.aj4tt" + tempCalendar.GetFiscalMonth().ToString("D2") + "+aj4ob" + tempCalendar.GetFiscalMonth().ToString("D2") + ") as balance from cmsdat.glmt as a, cmsdat.mast as b where a.aj4comp=49 and a.aj4ccyy=" + (tempCalendar.GetFiscalYear() + 2000).ToString("D2") + " and a.aj4gl#1=b.azgl#1 and a.aj4gl#2=b.azgl#2 and b.azcomp=49";
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                Account account = new Account();
                account.glNum1 = Convert.ToInt32(reader["no1"]) % 100;
                account.glNum2 = Convert.ToInt32(reader["no2"]);
                account.title = reader["title"].ToString().Trim();
                account.balanceAmount49 = Convert.ToDouble(reader["balance"]);
                int key = account.glNum1 * 1000000 + account.glNum2;
                if (99999999 == key)
                {
                    continue;
                }
                if (accountMap.ContainsKey(key))
                {
                    accountMap[key].balanceAmount49 = account.balanceAmount49;
                }
                else
                {
                    accountMap.Add(key, account);
                }
            }
            reader.Close();
        }
        else if (SelectMarkham.Checked)
        {
            query = "select distinct a.aj4gl#1 as no1, a.aj4gl#2 as no2, b.aztitl as title, b.azatyp as account_type, (a.aj4tt" + calendar.GetFiscalMonth().ToString("D2") + "+aj4ob" + calendar.GetFiscalMonth().ToString("D2") + ") as balance from cmsdat.glmt as a, cmsdat.mast as b where a.aj4comp=1 and a.aj4ccyy=" + (calendar.GetFiscalYear() + 2000).ToString("D2") + " and a.aj4gl#1=b.azgl#1 and a.aj4gl#2=b.azgl#2 and b.azcomp=1 and a.aj4gl#1!=200";
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                Account account = new Account();
                account.glNum1 = Convert.ToInt32(reader["no1"]) % 100;
                account.glNum2 = Convert.ToInt32(reader["no2"]);
                account.title = reader["title"].ToString().Trim();
                account.balanceAmount01 = Convert.ToDouble(reader["balance"]);
                int key = account.glNum1 * 1000000 + account.glNum2;
                if (99999999 == key)
                {
                    continue;
                }
                if (accountMap.ContainsKey(key))
                {
                    accountMap[key].balanceAmount01 = account.balanceAmount01;
                }
                else
                {
                    accountMap.Add(key, account);
                }
            }
            reader.Close();
        }
        else if (SelectMichigan.Checked)
        {
            query = "select a.aj4gl#1 as no1, a.aj4gl#2 as no2, b.aztitl as title, b.azatyp as account_type, (a.aj4tt" + calendar.GetFiscalMonth().ToString("D2") + "+aj4ob" + calendar.GetFiscalMonth().ToString("D2") + ") as balance from cmsdat.glmt as a, cmsdat.mast as b where a.aj4comp=3 and a.aj4ccyy=" + (calendar.GetFiscalYear() + 2000).ToString("D2") + " and a.aj4gl#1=b.azgl#1 and a.aj4gl#2=b.azgl#2 and b.azcomp=3";
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                Account account = new Account();
                account.glNum1 = Convert.ToInt32(reader["no1"]) % 100;
                account.glNum2 = Convert.ToInt32(reader["no2"]);
                account.title = reader["title"].ToString().Trim();
                account.balanceAmount03 = Convert.ToDouble(reader["balance"]);
                int key = account.glNum1 * 1000000 + account.glNum2;
                if (99999999 == key)
                {
                    continue;
                }
                if (accountMap.ContainsKey(key))
                {
                    accountMap[key].balanceAmount03 = account.balanceAmount03;
                }
                else
                {
                    accountMap.Add(key, account);
                }
            }
            reader.Close();
        }
        else if (SelectTexas.Checked)
        {
            query = "select a.aj4gl#1 as no1, a.aj4gl#2 as no2, b.aztitl as title, b.azatyp as account_type, (a.aj4tt" + calendar.GetFiscalMonth().ToString("D2") + "+aj4ob" + calendar.GetFiscalMonth().ToString("D2") + ") as balance from cmsdat.glmt as a, cmsdat.mast as b where a.aj4comp=5 and a.aj4ccyy=" + (calendar.GetFiscalYear() + 2000).ToString("D2") + " and a.aj4gl#1=b.azgl#1 and a.aj4gl#2=b.azgl#2 and b.azcomp=5";
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                Account account = new Account();
                account.glNum1 = Convert.ToInt32(reader["no1"]) % 100;
                account.glNum2 = Convert.ToInt32(reader["no2"]);
                account.title = reader["title"].ToString().Trim();
                account.balanceAmount05 = Convert.ToDouble(reader["balance"]);
                int key = account.glNum1 * 1000000 + account.glNum2;
                if (99999999 == key)
                {
                    continue;
                }
                if (accountMap.ContainsKey(key))
                {
                    accountMap[key].balanceAmount05 = account.balanceAmount05;
                }
                else
                {
                    accountMap.Add(key, account);
                }
            }
            reader.Close();
        }
        // write to data table
        var accountList = from account in accountMap.Values orderby account.glNum1 * 1000000 + account.glNum2 select account;
        foreach (var account in accountList)
        {
            DataRow row = dataTable.NewRow();
            // account # and name
            row[0] = account.glNum1.ToString("D2") + "-" + (account.glNum2 / 100).ToString("D4") + "-" + (account.glNum2 % 100).ToString("D2");
            row[1] = account.title;
            // amount
            if (1 == plantID)
            {
                row[2] = account.balanceAmount01.ToString("C2");
            }
            else if (3 == plantID)
            {
                row[2] = account.balanceAmount03.ToString("C2");
            }
            else if (5 == plantID)
            {
                row[2] = account.balanceAmount05.ToString("C2");
            }
            else if (4 == plantID)
            {
                row[2] = account.balanceAmount04.ToString("C2");
                row[3] = account.balanceAmount41.ToString("C2");
                row[4] = account.balanceAmount48.ToString("C2");
                row[5] = account.balanceAmount49.ToString("C2");
                row[6] = (account.balanceAmount04 + account.balanceAmount41 + account.balanceAmount48 + account.balanceAmount49).ToString("C2");
            }
            dataTable.Rows.Add(row);
        }
        dataTable.AcceptChanges();
    }

    private void GenerateCSVFile(string filePath)
    {
        // create file
        StreamWriter writer = new StreamWriter(filePath);
        // insert title
        writer.WriteLine("EXCO " + Session["_plantName"].ToString() + " Trial balance at " + DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss"));
        writer.WriteLine();
        foreach (TableCell cell in GridViewTrialBalance.HeaderRow.Cells)
        {
            writer.Write(cell.Text + ",");
        }
        writer.WriteLine();
        // write content
        foreach (GridViewRow row in GridViewTrialBalance.Rows)
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

    protected void ButtonSave_Click(object sender, EventArgs e)
    {
        if (GridViewTrialBalance.Rows.Count > 0)
        {
            // write to excel
            string fileName = @"C:\Trial Balance\" + User.Identity.Name + " Trial Balance for " + Session["_plantName"].ToString() + " at " + DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss") + ".csv";
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

    protected void ButtonGenerate_Click(object sender, EventArgs e)
    {
        GridViewTrialBalance.Visible = true;
        // build data table
        BuildDataTable();
        // fill data into data table
        FillDataTable();
        // fill grid view
        GridView gridView = GridViewTrialBalance;
        gridView.Caption = "<font size=\"9\" color=\"red\">Trial Balance Report for ";
        if (SelectMarkham.Checked)
        {
            Session["_plantName"] = "Markham";
            gridView.Caption += "Markham (CAD)";
        }
        if (SelectMichigan.Checked)
        {
            Session["_plantName"] = "Michigan";
            gridView.Caption += "Michigan (USD)";
        }
        if (SelectTexas.Checked)
        {
            Session["_plantName"] = "Texas";
            gridView.Caption += "Texas (USD)";
        }
        if (SelectColombia.Checked)
        {
            Session["_plantName"] = "Colombia";
            gridView.Caption += "Colombia (PESO)";
        }
        gridView.Caption += " at " + DateTime.Now.ToString("MM/dd/yyyy H:mm") + "</font>";
        gridView.DataSource = dataTable;
        gridView.DataBind();
    }
}