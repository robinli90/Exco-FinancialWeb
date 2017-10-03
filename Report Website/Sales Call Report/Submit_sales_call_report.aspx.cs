using System;
using System.Net;
using System.IO;
using System.Net.Mail;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.Odbc;
using ExcoUtility;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using PdfSharp.Drawing;

public partial class submit_sales_call_report : System.Web.UI.Page
{
    private bool isEdit = false;

    private string filesubPath = String.Empty;

    protected void Page_Init(object sender, EventArgs e)
    {
        if (0 == Session["_SalesCallReportID"].ToString().CompareTo("0"))
        {
            DateCalendar.Visible = true;
            DateCalendar.SelectionChanged += new EventHandler(Date_Select);
            FollowUpCalendar.Visible = true;
            FollowUpCalendar.SelectionChanged += new EventHandler(FollowUpDate_Select);
            CustomerText.Visible = false;
            SalesLabel.Visible = false;
            SalesText.Visible = false;
            SubmitButton.Visible = true;
            ClearButton.Visible = true;
        }
        // display legacy report detail
        else
        {
            isEdit = true;
            SalesLabel.Visible = true;
            SalesText.Visible = true;
            FollowUpCalendar.Visible = false;
            DateCalendar.Visible = false;


            string query = "select [CustomerName], [Sales], [Date], [Attendee], [NumberOfPresses], [DieExpense], [OtherSuppliers], [PurposeOfVisit], [AdditionalInformation], [HowIncreaseBusiness], [FollowUpDate] from [tiger].[dbo].[Sales_Call_Report] where [ID]=" + Session["_SalesCallReportID"];
            ExcoODBC database = ExcoODBC.Instance;
            database.Open(Database.DECADE_MARKHAM);
            OdbcDataReader reader = database.RunQuery(query);
            if (reader.Read())
            {
                CustomerList.SelectedValue = reader[0].ToString();
                CustomerText.Text = reader[0].ToString();
                SalesText.Text = reader[1].ToString();
                DateText.Text = reader[2].ToString();
                AttendeeText.Text = reader[3].ToString();
                PressText.Text = reader[4].ToString();
                ExpenseText.Text = reader[5].ToString();
                SupplierText.Text = reader[6].ToString();
                PurposeText.Text = reader[7].ToString();
                InformationText.Text = reader[8].ToString();
                IncreaseText.Text = reader[9].ToString();
                FollowUpText.Text = reader[10].ToString();

                if (SalesText.Text != User.Identity.Name)
                {
                    CustomerText.Visible = true;
                    CustomerList.Visible = false;
                    CustomerText.ReadOnly = true;
                    SalesText.ReadOnly = true;
                    DateText.ReadOnly = true;
                    FollowUpText.ReadOnly = true;
                    AttendeeText.ReadOnly = true;
                    PressText.ReadOnly = true;
                    ExpenseText.ReadOnly = true;
                    SupplierText.ReadOnly = true;
                    PurposeText.ReadOnly = true;
                    InformationText.ReadOnly = true;
                    IncreaseText.ReadOnly = true;
                    SubmitButton.Visible = false;
                    ClearButton.Visible = false;
                }
                else
                {
                    CustomerText.Visible = false;
                    CustomerList.Visible = true;
                    SubmitButton.Visible = true;
                    ClearButton.Visible = true;
                }
            }
            else
            {
                Response.Write("<SCRIPT LANGUAGE=\"JavaScript\">alert(\"Invalid report!\")</SCRIPT>");
            }
            reader.Close();
        }
    }


    protected void Page_Load(object sender, EventArgs e)
    {

    }

    // write to DateText when pick up a date
    protected void Date_Select(object sender, EventArgs e)
    {
        DateText.Text = DateCalendar.SelectedDate.ToShortDateString();
    }

    // write to FollowUpText when pick up a date
    protected void FollowUpDate_Select(object sender, EventArgs e)
    {
        FollowUpText.Text = FollowUpCalendar.SelectedDate.ToShortDateString();
    }

    // submit sales call report
    protected void ClearButton_Click(object sender, EventArgs e)
    {
        // clear text
        CustomerList.SelectedIndex = 0;
        DateText.Text = string.Empty;
        AttendeeText.Text = string.Empty;
        PressText.Text = string.Empty;
        ExpenseText.Text = string.Empty;
        SupplierText.Text = string.Empty;
        PurposeText.Text = string.Empty;
        InformationText.Text = string.Empty;
        IncreaseText.Text = string.Empty;
        FollowUpText.Text = string.Empty;
        isEdit = false;
        // reset calendar
        DateCalendar.SelectedDate = DateTime.Today;
        FollowUpCalendar.SelectedDate = DateTime.Today;
    }

    // clear all typed text
    protected void SubmitButton_Click(object sender, EventArgs e)
    {
        // check if mandatory fields are filled and if data is valid
        if (DateText.Text.Length == 0)
        {
            Response.Write("<SCRIPT LANGUAGE=\"JavaScript\">alert(\"Please input visit date!\")</SCRIPT>");
        }
        // write to decade
        else
        {
            /* validation
            // check if type is ok
            try
            {               
                Convert.ToInt32(PressText.Text);
            }
            catch
            {
                Response.Write("<SCRIPT LANGUAGE=\"JavaScript\">alert(\"Please check number of presses!\")</SCRIPT>");
                return;
            }

            try
            {
                Convert.ToDouble(ExpenseText.Text);
            }
            catch
            {
                Response.Write("<SCRIPT LANGUAGE=\"JavaScript\">alert(\"Please check Estimated Die Expense/Mn!\")</SCRIPT>");
                return;
            }
            */

            // get customer id
            string query = "select bvcust from cmsdat.cust where bvname='" + CustomerList.Text + "' order by bvcust";
            string customerID = string.Empty;
            ExcoODBC database = ExcoODBC.Instance;
            database.Open(Database.CMSDAT);
            OdbcDataReader reader = database.RunQuery(query);
            if (reader.Read())
            {
                customerID = reader["bvcust"].ToString().Trim();
            }
            else
            {
                throw new Exception("Getting customer id failed!");
            }
            reader.Close();
            // write to database
            if (isEdit)
            {
                query = "update [tiger].[dbo].[Sales_Call_Report] set [CustomerName] = '" + CustomerList.Text.Replace("'", "''") + "', [CustomerID] = '" + customerID + "', [Date] = '" + DateText.Text + "', [Attendee] = '" + AttendeeText.Text.Replace("'", "''") + "', [NumberOfPresses] = '" + PressText.Text.Replace("'", "''") + "', [DieExpense] = '" + ExpenseText.Text.Replace("'", "''") + "', [OtherSuppliers] = '" + SupplierText.Text.Replace("'", "''") + "', [PurposeOfVisit] = '" + PurposeText.Text.Replace("'", "''") + "', [AdditionalInformation] = '" + InformationText.Text.Replace("'", "''") + "', [HowIncreaseBusiness] = '" + IncreaseText.Text.Replace("'", "''") + "', [FollowUpDate] = '" + FollowUpText.Text.Replace("'", "''") + "' where [ID]= '" + Session["_SalesCallReportID"] + "'";
            }
            else
            {               
                query = "insert into [tiger].[dbo].[Sales_Call_Report] values ('" + CustomerList.Text.Replace("'", "''") + "', '" + customerID + "', '" + User.Identity.Name + "', '" + DateText.Text + "', '" + AttendeeText.Text + "', '" + PressText.Text.Replace("'", "''") + "', '" + ExpenseText.Text + "', '" + SupplierText.Text.Replace("'", "''") + "', '" + PurposeText.Text.Replace("'", "''") + "', '" + InformationText.Text.Replace("'", "''") + "', '" + IncreaseText.Text.Replace("'", "''") + "', '" + FollowUpText.Text + "')";
            }
            database.Open(Database.DECADE_MARKHAM);
            if (1 == database.RunQueryWithoutReader(query))
            {
                // generate pdf file
                string filePath = GeneratePDFFile(customerID);

                // send to recipients
                //SendFileViaEmail(filePath);
                //System.Diagnostics.Process.Start(filePath);
                //Response.Write("<SCRIPT LANGUAGE=\"JavaScript\">alert(\"Submit report done!\")</SCRIPT>");                        
                Response.Redirect("EmailPDF.aspx?filesubPath=" + filesubPath);
            }
            else
            {
                Response.Write("<SCRIPT LANGUAGE=\"JavaScript\">alert(\"Submit report failed!\")</SCRIPT>");
            }            
        }
        isEdit = false;
    }

    // generate pdf file
    private string GeneratePDFFile(string customerID)
    {     
        Document document = new Document();
        document.DefaultPageSetup.TopMargin = "1cm";
        document.DefaultPageSetup.LeftMargin = "1cm";
        document.DefaultPageSetup.RightMargin = "1cm";
        document.DefaultPageSetup.BottomMargin = "1cm";
        Section section = document.AddSection();
        // insert logo
        MigraDoc.DocumentObjectModel.Shapes.Image logo = section.AddImage(@"C:\Sales Call Report\Exco_logo.png");
        //MigraDoc.DocumentObjectModel.Shapes.Image logo = section.AddImage(@"\\10.0.0.6\Shopdata\Development\tiger\Exco_logo.png");
        
        logo.LockAspectRatio = true;
        logo.ScaleHeight = 0.5;
        section.AddParagraph().AddLineBreak();
        // add title
        Paragraph paragraph = section.AddParagraph("Sales Call Report");
        paragraph.Format.Font.Color = Colors.Black;
        paragraph.Format.Font.Bold = true;
        paragraph.Format.Font.Size = "30";
        paragraph.Format.Alignment = ParagraphAlignment.Center;
        section.AddParagraph(DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss")).Format.Alignment = ParagraphAlignment.Center; ;
        // add customer information
        string query = "select bvadr1, bvadr2, bvadr3, bvadr4, bvtelp, bvfax, bvcont, bvemal from cmsdat.cust where bvcust like '" + customerID + "'";
        string customerAdd1 = string.Empty;
        string customerAdd2 = string.Empty;
        string customerAdd3 = string.Empty;
        string customerAdd4 = string.Empty;
        string customerTel = string.Empty;
        string customerFax = string.Empty;
        string customerContact = string.Empty;
        string customerEmail = string.Empty;
        ExcoODBC database = ExcoODBC.Instance;
        database.Open(Database.CMSDAT);
        OdbcDataReader reader = database.RunQuery(query);
        if (reader.Read())
        {
            customerAdd1 = reader["bvadr1"].ToString().Trim();
            customerAdd2 = reader["bvadr2"].ToString().Trim();
            customerAdd3 = reader["bvadr3"].ToString().Trim();
            customerAdd4 = reader["bvadr4"].ToString().Trim();
            customerTel = reader["bvtelp"].ToString().Trim();
            customerFax = reader["bvfax"].ToString().Trim();
            customerContact = reader["bvcont"].ToString().Trim();
            customerEmail = reader["bvemal"].ToString().Trim().Replace("mailto:", "");
        }
        else
        {
            throw new Exception("Getting customer information failed!");
        }
        reader.Close();
        paragraph = section.AddParagraph();
        paragraph.AddText(CustomerList.Text);
        paragraph.AddLineBreak();
        if (customerContact.Length > 0)
        {
            paragraph.AddText("Contact Person: " + customerContact);
            paragraph.AddLineBreak();
        }
        if (customerAdd1.Length > 0)
        {
            paragraph.AddText(customerAdd1);
            paragraph.AddLineBreak();
        }
        if (customerAdd2.Length > 0)
        {
            paragraph.AddText(customerAdd2);
            paragraph.AddLineBreak();
        }
        if (customerAdd3.Length > 0)
        {
            paragraph.AddText(customerAdd3);
            paragraph.AddLineBreak();
        }
        if (customerAdd4.Length > 0)
        {
            paragraph.AddText(customerAdd4);
            paragraph.AddLineBreak();
        }
        if (customerTel.Length > 0)
        {
            paragraph.AddText("Tel: " + customerTel);
            paragraph.AddLineBreak();
        }
        if (customerFax.Length > 0)
        {
            paragraph.AddText("Fax: " + customerFax);
            paragraph.AddLineBreak();
        }
        if (customerEmail.Length > 0)
        {
            paragraph.AddText("Email: " + customerEmail);
            paragraph.AddLineBreak();
        }
        // add sales person information
        query = "select Email, FirstName, LastName from tiger.dbo.[user] where UserName like '" + User.Identity.Name + "'";
        string salesEmail = string.Empty;
        string salesName = string.Empty;
        database.Open(Database.DECADE_MARKHAM);
        reader = database.RunQuery(query);
        if (reader.Read())
        {
            salesEmail = reader["Email"].ToString().Trim();
            salesName = reader["FirstName"].ToString().Trim() + " " + reader["LastName"].ToString().Trim();
        }
        else
        {
            throw new Exception("Getting sales information failed!");
        }
        reader.Close();
        // add sales call report information
        section.AddParagraph().AddLineBreak();
        MigraDoc.DocumentObjectModel.Tables.Table table = section.AddTable();
        table.Format.Alignment = ParagraphAlignment.Center;
        table.Style = "Table";
        table.Borders.Width = 1;
        table.Borders.Color = Colors.Black;
        table.Borders.Width = 0.25;
        table.Borders.Left.Width = 0.5;
        table.Borders.Right.Width = 0.5;
        table.Rows.LeftIndent = 0;
        MigraDoc.DocumentObjectModel.Tables.Column column = table.AddColumn("3.5cm");
        column.Format.Alignment = ParagraphAlignment.Center;
        column = table.AddColumn("3cm");
        column.Format.Alignment = ParagraphAlignment.Center;
        column = table.AddColumn("3cm");
        column.Format.Alignment = ParagraphAlignment.Center;
        column = table.AddColumn("3cm");
        column.Format.Alignment = ParagraphAlignment.Center;
        column = table.AddColumn("3cm");
        column.Format.Alignment = ParagraphAlignment.Center;
        column = table.AddColumn("3cm");
        column.Format.Alignment = ParagraphAlignment.Center;
        MigraDoc.DocumentObjectModel.Tables.Row row = table.AddRow();
        row.Cells[0].AddParagraph("Sales");
        row.Cells[0].Format.Font.Bold = true;
        row.Cells[1].MergeRight = 1;
        row.Cells[1].AddParagraph(salesName);
        row.Cells[3].AddParagraph("Email");
        row.Cells[3].Format.Font.Bold = true;
        row.Cells[4].MergeRight = 1;
        row.Cells[4].AddParagraph(salesEmail);
        row = table.AddRow();
        row.Cells[0].AddParagraph("Visit");
        row.Cells[0].Format.Font.Bold = true;
        row.Cells[1].MergeRight = 1;
        row.Cells[1].AddParagraph(DateText.Text);
        row.Cells[3].AddParagraph("Follow Up");
        row.Cells[3].Format.Font.Bold = true;
        row.Cells[4].MergeRight = 1;
        row.Cells[4].AddParagraph(FollowUpText.Text);     
        row = table.AddRow();
        row.Cells[0].AddParagraph("Number of Presses");
        row.Cells[0].Format.Font.Bold = true;
        row.Cells[1].AddParagraph(PressText.Text);
        row.Cells[2].AddParagraph("Est. Die Exp./Mn");
        row.Cells[2].Format.Font.Bold = true;
        row.Cells[3].AddParagraph(ExpenseText.Text);
        row.Cells[4].AddParagraph("Other Suppliers");
        row.Cells[4].Format.Font.Bold = true;
        row.Cells[5].AddParagraph(SupplierText.Text);
        row = table.AddRow();
        row.Cells[0].AddParagraph("Attendee's");
        row.Cells[0].Format.Font.Bold = true;
        row.Cells[1].MergeRight = 4;
        row.Cells[1].AddParagraph(AttendeeText.Text);   
        row = table.AddRow();
        row.Cells[0].AddParagraph("Purpose of Visit");
        row.Cells[0].Format.Font.Bold = true;
        row.Cells[1].MergeRight = 4;
        row.Cells[1].AddParagraph(AdjustWordWrap(PurposeText.Text));
        row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
        row = table.AddRow();
        row.Cells[0].AddParagraph("Addtional Information");
        row.Cells[0].Format.Font.Bold = true;
        row.Cells[1].MergeRight = 4;
        row.Cells[1].AddParagraph(AdjustWordWrap(InformationText.Text));
        row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
        row = table.AddRow();
        row.Cells[0].AddParagraph("How can ETS Increase Business");
        row.Cells[0].Format.Font.Bold = true;
        row.Cells[1].MergeRight = 4;
        row.Cells[1].AddParagraph(AdjustWordWrap(IncreaseText.Text));
        row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
        // render pdf
        PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(false, PdfFontEmbedding.Always);
        pdfRenderer.Document = document;
        pdfRenderer.RenderDocument();
        filesubPath = @"sales call report by " + User.Identity.Name + " at " + DateTime.Now.ToString("yyyy-MM-dd HH.mm.ss") + ".pdf";
        //string filePath = @"C:\Sales Call Report\" + filesubPath;  
        //string filePath = @"\\10.0.0.6\Shopdata\Development\tiger\Sales Call Report\sales call report by " + User.Identity.Name + " at " + DateTime.Now.ToString("yyyy-MM-dd HH.mm.ss") + ".pdf";                      
        //string filePath = @"C:\Sales Call Report\" + filesubPath;
        string filePath = Server.MapPath("~/Sales Call Report PDF/" + filesubPath);
        if (Directory.Exists(Server.MapPath("~/Sales Call Report PDF/")))
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
        else
        {
            Directory.CreateDirectory(Server.MapPath("~/Sales Call Report PDF/"));
        }

        pdfRenderer.PdfDocument.Save(filePath);

        return filePath;
    }

    // adjust word wrap
    private string AdjustWordWrap(string text)
    {
        string output = text;
        int index = 75;
        while (index < text.Length)
        {
            while (output[index] != ' ')
            {
                index -= 1;
            }
            output = output.Insert(index, "\xad");
            index += 76;
        }
        return output;
    }

    // send pdf file to every recipients
    private void SendFileViaEmail(string filePath)
    {
        // get recipients list
        List<string> recipientList = new List<string>();
        string query = "select Email from tiger.dbo.[Sales_Call_Report_Recipients]";
        ExcoODBC database = ExcoODBC.Instance;
        database.Open(Database.DECADE_MARKHAM);
        OdbcDataReader reader = database.RunQuery(query);
        while (reader.Read())
        {
            recipientList.Add(reader["Email"].ToString().Trim());
        }
        reader.Close();
        // send file
        MailMessage mailmsg = new MailMessage();
        foreach (string recipient in recipientList)
        {
            mailmsg.To.Add(recipient);
        }
        mailmsg.From = new MailAddress("office@etsdies.com");
        mailmsg.Subject = "Sales Call Report";
        mailmsg.Attachments.Add(new Attachment(filePath));
        // smtp client
        SmtpClient client = new SmtpClient("smtp.pathcom.com", 587);
        //NetworkCredential credential = new NetworkCredential("report@etsdies.com", "5Zh2P8k4");
        NetworkCredential credential = new NetworkCredential("office@etsdies.com", "F130Desk");
        client.Credentials = credential;
        bool hasSend = false;
        while (!hasSend)
        {
            try
            {
                client.Send(mailmsg);
                hasSend = true;
            }
            catch
            {
                hasSend = false;
            }
        }
    }
}