using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data.Odbc;
using System.ServiceModel.Security;

public partial class SiteMaster : System.Web.UI.MasterPage
{
    // priviledge list
    // sales call report
    private const ulong hasSalesCallReport = 1 << 0;
    // invoice register
    private const ulong hasInvoiceRegisterColombia = 1 << 1;
    private const ulong hasInvoiceRegisterTexas = 1 << 2;
    private const ulong hasInvoiceRegisterMichigan = 1 << 3;
    private const ulong hasInvoiceRegisterMarkham = 1 << 4;
    // Trial balance
    private const ulong hasTrialBalanceMarkham = 1 << 5;
    private const ulong hasTrialBalanceMichigan = 1 << 6;
    private const ulong hasTrialBalanceTexas = 1 << 7;
    private const ulong hasTrialBalanceColombia = 1 << 8;
    // sales report
    private const ulong hasSalesReportMarkham = 1 << 9;
    private const ulong hasSalesReportMichigan = 1 << 10;
    private const ulong hasSalesReportTexas = 1 << 11;
    private const ulong hasSalesReportColombia = 1 << 12;
    // income statement report
    private const ulong hasIncomeStatementReportMarkham = 1 << 13;
    private const ulong hasIncomeStatementReportMichigan = 1 << 14;
    private const ulong hasIncomeStatementReportTexas = 1 << 15;
    private const ulong hasIncomeStatementReportColombia = 1 << 16;
    // monthly sales by customer report
    private const ulong hasMonthlySalesByCustomer = 1 << 17;

    protected void Page_Load(object sender, EventArgs e)
    {
        GetUserPrivilege(HttpContext.Current.User.Identity.Name);
        if (HttpContext.Current.User.Identity.IsAuthenticated)
        {
            NavigationMenu.Visible = true;
            // sales call report
            if (true == (bool)Session["_hasSalesCallReport"])
            {
                NavigationMenu.Items.Add(new MenuItem("Sales Call Report", "sales_call_report", "", "Sales Call Report/sales_call_report.aspx"));
            }
            // invoice register
            if (true == (bool)Session["_hasInvoiceRegisterTexasReport"] || true == (bool)Session["_hasInvoiceRegisterColombiaReport"] || true == (bool)Session["_hasInvoiceRegisterMichiganReport"] || true == (bool)Session["_hasInvoiceRegisterMarkhamReport"])
            {
                NavigationMenu.Items.Add(new MenuItem("Invoice Register", "invoice_register", "", "Invoice Register/invoice_register.aspx"));
            }
            // sales report
            if (true == (bool)Session["_hasSalesReportMarkham"] || true == (bool)Session["_hasSalesReportMichigan"] || true == (bool)Session["_hasSalesReportTexas"] || true == (bool)Session["_hasSalesReportColombia"])
            {
                NavigationMenu.Items.Add(new MenuItem("Sales Report", "sales_report", "", "Sales Report/sales_report.aspx"));
            }
            else if (HttpContext.Current.User.Identity.Name == "JYOUNG")
            {
                NavigationMenu.Items.Add(new MenuItem("Sales Report", "sales_report", "", "Sales Report/sales_report.aspx"));
            }
            // trial balance
            if (true == (bool)Session["_hasTrialBalanceMarkham"] || true == (bool)Session["_hasTrialBalanceMichigan"] || true == (bool)Session["_hasTrialBalanceTexas"] || true == (bool)Session["_hasTrialBalanceColombia"])
            {
                NavigationMenu.Items.Add(new MenuItem("Trial Balance", "Trial_balance", "", "Trial Balance/Trial_balance.aspx"));
            }
            // income statement report
            if (true == (bool)Session["_hasIncomeStatementReportMarkham"] || true == (bool)Session["_hasIncomeStatementReportMichigan"] || true == (bool)Session["_hasIncomeStatementReportTexas"] || true == (bool)Session["_hasIncomeStatementReportColombia"])
            {
                NavigationMenu.Items.Add(new MenuItem("Income Statement Report", "income_statement_report", "", "Income Statement Report/income_statement_report.aspx"));
            }
            // monthly sales by customer report
            if (true == (bool)Session["_hasMonthlySalesByCustomer"])
            {
                NavigationMenu.Items.Add(new MenuItem("Sapa Sales by Month", "monthly_sales_by_customer", "", "Monthly Sales By Customer/monthly_sales_by_customer.aspx"));
                NavigationMenu.Items.Add(new MenuItem("Customer Die Purchases", "invoice_register", "", "invoice register/Copy of invoice_register.aspx"));
                NavigationMenu.Items.Add(new MenuItem("Vendor Parts Listing", "invoice_register", "", "invoice register/PO.aspx"));
                NavigationMenu.Items.Add(new MenuItem("Customer Purchases", "invoice_register", "", "Customer Lookup/CustPurchases.aspx"));
                //NavigationMenu.Items.Add(new MenuItem("test", "test", "", "test.aspx"));
                //NavigationMenu.Items.Add(new MenuItem("Income Statement Report", "income_statement_report", "", "Income Statement Report/income_statement_report.aspx"));
            }
        }
        else
        {
            NavigationMenu.Visible = false;
        }
    }

    private void GetUserPrivilege(string userName)
    {
        string strConnection = "server=10.0.0.6;database=tiger;user id=jamie;password=jamie;";
        SqlConnection sqlConnection = new SqlConnection(strConnection);
        sqlConnection.Open();
        String SQLQuery = "select [Privilege] from [tiger].[dbo].[user] where [UserName] like '" + userName + "'";
        SqlCommand command = new SqlCommand(SQLQuery, sqlConnection);
        SqlDataReader reader;
        reader = command.ExecuteReader();
        ulong privilege = 0;
        if (reader.Read())
        {
            privilege = Convert.ToUInt64(reader[0]);
        }
        reader.Close();
        // get user priviledge
        // sales call report
        if (!(userName == "LSILVA") && !(userName == "YGOMEZ"))
        {
            if ((privilege & hasSalesCallReport) != ulong.MinValue)
            {
                Session["_hasSalesCallReport"] = true;
            }
            else
            {
                Session["_hasSalesCallReport"] = false;
            }
            // invoice register
            if ((privilege & hasInvoiceRegisterColombia) != ulong.MinValue)
            {
                Session["_hasInvoiceRegisterColombiaReport"] = true;
            }
            else
            {
                Session["_hasInvoiceRegisterColombiaReport"] = false;
            }
            if ((privilege & hasInvoiceRegisterTexas) != ulong.MinValue)
            {
                Session["_hasInvoiceRegisterTexasReport"] = true;
            }
            else
            {
                Session["_hasInvoiceRegisterTexasReport"] = false;
            }
            if ((privilege & hasInvoiceRegisterMichigan) != ulong.MinValue)
            {
                Session["_hasInvoiceRegisterMichiganReport"] = true;
            }
            else
            {
                Session["_hasInvoiceRegisterMichiganReport"] = false;
            }
            if ((privilege & hasInvoiceRegisterMarkham) != ulong.MinValue)
            {
                Session["_hasInvoiceRegisterMarkhamReport"] = true;
            }
            else
            {
                Session["_hasInvoiceRegisterMarkhamReport"] = false;
            }
            // sales report
            if (((privilege & hasSalesReportMarkham) != ulong.MinValue) || HttpContext.Current.User.Identity.Name == "JYOUNG")
            {
                Session["_hasSalesReportMarkham"] = true;
            }
            else
            {
                Session["_hasSalesReportMarkham"] = false;
            }
            if (((privilege & hasSalesReportMichigan) != ulong.MinValue) || HttpContext.Current.User.Identity.Name == "JYOUNG")
            {
                Session["_hasSalesReportMichigan"] = true;
            }
            else
            {
                Session["_hasSalesReportMichigan"] = false;
            }
            if (((privilege & hasSalesReportTexas) != ulong.MinValue) || HttpContext.Current.User.Identity.Name == "YGOMEZ" || HttpContext.Current.User.Identity.Name == "JYOUNG")
            {
                Session["_hasSalesReportTexas"] = true;
            }
            else
            {
                Session["_hasSalesReportTexas"] = false;
            }
            if (((privilege & hasSalesReportColombia) != ulong.MinValue) || HttpContext.Current.User.Identity.Name == "JYOUNG")
            {
                Session["_hasSalesReportColombia"] = true;
            }
            else
            {
                Session["_hasSalesReportColombia"] = false;
            }
            // trial balance
            if ((privilege & hasTrialBalanceColombia) != ulong.MinValue)
            {
                Session["_hasTrialBalanceColombia"] = true;
            }
            else
            {
                Session["_hasTrialBalanceColombia"] = false;
            }
            if ((privilege & hasTrialBalanceTexas) != ulong.MinValue)
            {
                Session["_hasTrialBalanceTexas"] = true;
            }
            else
            {
                Session["_hasTrialBalanceTexas"] = false;
            }
            if ((privilege & hasTrialBalanceMichigan) != ulong.MinValue)
            {
                Session["_hasTrialBalanceMichigan"] = true;
            }
            else
            {
                Session["_hasTrialBalanceMichigan"] = false;
            }
            if ((privilege & hasTrialBalanceMarkham) != ulong.MinValue)
            {
                Session["_hasTrialBalanceMarkham"] = true;
            }
            else
            {
                Session["_hasTrialBalanceMarkham"] = false;
            }
            // income statement report
            if ((privilege & hasIncomeStatementReportColombia) != ulong.MinValue)
            {
                Session["_hasIncomeStatementReportColombia"] = true;
            }
            else
            {
                Session["_hasIncomeStatementReportColombia"] = false;
            }
            if ((privilege & hasIncomeStatementReportTexas) != ulong.MinValue)
            {
                Session["_hasIncomeStatementReportTexas"] = true;
            }
            else
            {
                Session["_hasIncomeStatementReportTexas"] = false;
            }
            if ((privilege & hasIncomeStatementReportMichigan) != ulong.MinValue)
            {
                Session["_hasIncomeStatementReportMichigan"] = true;
            }
            else
            {
                Session["_hasIncomeStatementReportMichigan"] = false;
            }
            if ((privilege & hasIncomeStatementReportMarkham) != ulong.MinValue)
            {
                Session["_hasIncomeStatementReportMarkham"] = true;
            }
            else
            {
                Session["_hasIncomeStatementReportMarkham"] = false;
            }
            // monthly sales by customer report
            if ((privilege & hasMonthlySalesByCustomer) != ulong.MinValue)
            {
                Session["_hasMonthlySalesByCustomer"] = true;
            }
            else
            {
                Session["_hasMonthlySalesByCustomer"] = false;
            }

            if (HttpContext.Current.User.Identity.Name == "MROBBINS") // IF Matt
            // income statement report only for him
            {
                Session["_hasIncomeStatementReportColombia"] = true;
                Session["_hasIncomeStatementReportTexas"] = true;
                Session["_hasIncomeStatementReportMichigan"] = true;
                Session["_hasIncomeStatementReportMarkham"] = true;
                Session["_hasSalesCallReport"] = false;
                Session["_hasSalesReportMarkham"] = false;
                Session["_hasSalesReportMichigan"] = false;
                Session["_hasSalesReportTexas"] = false;
                Session["_hasSalesReportColombia"] = false;
            }
        }
        else if (userName == "LSILVA")
        {
            Session["_hasSalesCallReport"] = true;
            Session["_hasInvoiceRegisterColombiaReport"] = true;
            Session["_hasInvoiceRegisterTexasReport"] = false;
            Session["_hasInvoiceRegisterMichiganReport"] = false;
            Session["_hasInvoiceRegisterMarkhamReport"] = false;
            Session["_hasSalesReportMarkham"] = false;
            Session["_hasSalesReportMichigan"] = false;
            Session["_hasSalesReportTexas"] = false;
            Session["_hasSalesReportColombia"] = true;
            Session["_hasTrialBalanceColombia"] = true;
            Session["_hasTrialBalanceTexas"] = false;
            Session["_hasTrialBalanceMichigan"] = false;
            Session["_hasTrialBalanceMarkham"] = false;
            Session["_hasIncomeStatementReportColombia"] = true;
            Session["_hasIncomeStatementReportTexas"] = false;
            Session["_hasIncomeStatementReportMichigan"] = false;
            Session["_hasIncomeStatementReportMarkham"] = false;
            Session["_hasMonthlySalesByCustomer"] = true;
            //Session["_hasMonthlySalesByCustomer"] = false;
        }
        else if (userName == "YGOMEZ")
        {
            Session["_hasSalesCallReport"] = true;
            Session["_hasInvoiceRegisterColombiaReport"] = false;
            Session["_hasInvoiceRegisterTexasReport"] = true;
            Session["_hasInvoiceRegisterMichiganReport"] = false;
            Session["_hasInvoiceRegisterMarkhamReport"] = false;
            Session["_hasSalesReportMarkham"] = false;
            Session["_hasSalesReportMichigan"] = false;
            Session["_hasSalesReportTexas"] = true;
            Session["_hasSalesReportColombia"] = false;
            Session["_hasTrialBalanceColombia"] = false;
            Session["_hasTrialBalanceTexas"] = true;
            Session["_hasTrialBalanceMichigan"] = false;
            Session["_hasTrialBalanceMarkham"] = false;
            Session["_hasIncomeStatementReportColombia"] = false;
            Session["_hasIncomeStatementReportTexas"] = true;
            Session["_hasIncomeStatementReportMichigan"] = false;
            Session["_hasIncomeStatementReportMarkham"] = false;
            Session["_hasMonthlySalesByCustomer"] = false;
            //Session["_hasMonthlySalesByCustomer"] = false;
        }
    }
}