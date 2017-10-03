using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ExcoUtility;
using System.Data.Odbc;
using System.Drawing;
using System.Net.Mail;
using System.Net;
using System.IO;

public partial class Sales_Call_Report_EmailPDF : System.Web.UI.Page
{
    Panel pnlEmailBox;
    string filesubPath = String.Empty;
    string emailmsg = String.Empty;
    protected void Page_Init(object sender, EventArgs e)
    {
        filesubPath = Request.QueryString["filesubPath"];        
        //System.Diagnostics.Process.Start(filePath);     

        string query = "select * from [tiger].[dbo].[Sales_Call_Report_Recipients]";
        ExcoODBC database = ExcoODBC.Instance;
        database.Open(Database.DECADE_MARKHAM);
        OdbcDataReader reader = database.RunQuery(query);
        var emaillist = new Dictionary<string, string>();

        while (reader.HasRows)
        {
            while (reader.Read())
            {
                emaillist.Add(reader[0].ToString(), reader[1].ToString());
            }
            reader.NextResult();
        }
        reader.Close();
        DropDownEmailList.DataSource = emaillist;
        DropDownEmailList.DataTextField = "Key";
        DropDownEmailList.DataValueField = "Value";
        DropDownEmailList.DataBind();    

        //Get ContentPlaceHolder
        ContentPlaceHolder content = (ContentPlaceHolder)this.Master.FindControl("MainContent");

        Literal lt;
        Label lb;

        //Dynamic TextBox Panel
        pnlEmailBox = new Panel();
        pnlEmailBox.ID = "pnlEmailBox";
        //pnlEmailBox.BorderWidth = 1;
        pnlEmailBox.Width = 500;
        content.Controls.Add(pnlEmailBox);
        lt = new Literal();
        lt.Text = "<br />";
        content.Controls.Add(lt);
        lb = new Label();
        lb.Text = "Recipients List<br />";
        lb.Font.Size = FontUnit.Point(20);
        lb.ForeColor = Color.Black;

        pnlEmailBox.Controls.Add(lb);

        if (IsPostBack)
        {
            RecreateControls("RecipientList", "TextBox");
        }

        //Button To add TextBoxes
        Button btnAddTxt = new Button();
        btnAddTxt.ID = "btnAddTxt";
        btnAddTxt.Text = "Add Recipient";
        btnAddTxt.Click += new System.EventHandler(btnAdd_Click);
        content.Controls.Add(btnAddTxt);

        lt = new Literal();
        lt.Text = "<br /><br /><br /><br />";
        content.Controls.Add(lt);

        lb = new Label();
        lb.Text = "Email Message<br />";
        lb.Font.Size = FontUnit.Point(20);
        lb.ForeColor = Color.Black;
        content.Controls.Add(lb);

        TextBox txt = new TextBox();
        txt.ID = "EmailBody";
        txt.Width = 750;
        txt.Height = 300;
        txt.Wrap = true;
        txt.TextMode = TextBoxMode.MultiLine;
        if (IsPostBack)
        {
            string[] ctrls = Request.Form.ToString().Split('&');
            int cnt = FindOccurence("EmailBody");
            if (cnt > 0)
            {
                for (int i = 0; i < ctrls.Length; i++)
                {
                    if (ctrls[i].Contains("EmailBody"))
                    {
                        string ctrlName = ctrls[i].Split('=')[0];
                        string ctrlValue = ctrls[i].Split('=')[1];
                        ctrlValue = Server.UrlDecode(ctrlValue);   
                        txt.Text = ctrlValue;
                        emailmsg = ctrlValue;
                    }
                }
            }
        }
        content.Controls.Add(txt);

        lt = new Literal();
        lt.Text = "<br />";
        content.Controls.Add(lt);

        //Dummy Button To do PostBack
        Button Send = new Button();
        Send.ID = "btnSend";
        Send.Text = "Send";
        Send.Click += new System.EventHandler(btnSend_Click);
        content.Controls.Add(Send);
    }

    protected void btnAdd_Click(object sender, EventArgs e)
    {
        Button btn = (Button)sender;
        int cnt = FindOccurence("RecipientList");
        TextBox txt = new TextBox();
        txt.ID = "RecipientList-" + Convert.ToString(cnt + 1);
        pnlEmailBox.Controls.Add(txt);

        Literal lt = new Literal();
        lt.Text = "<br />";
        pnlEmailBox.Controls.Add(lt);

    }

    protected void btnSend_Click(object sender, EventArgs e)
    {
        // get recipients list
        List<string> recipientList = new List<string>();

        string[] ctrls = Request.Form.ToString().Split('&');
        int cnt = FindOccurence("RecipientList");
        if (cnt > 0)
        {
            for (int k = 1; k <= cnt; k++)
            {
                for (int i = 0; i < ctrls.Length; i++)
                {
                    if (ctrls[i].Contains("RecipientList" + "-" + k.ToString()))
                    {
                        string ctrlName = ctrls[i].Split('=')[0];
                        string ctrlValue = ctrls[i].Split('=')[1];
                        //Response.Write("<SCRIPT LANGUAGE=\"JavaScript\">alert("+ctrlValue+")</SCRIPT>");
                        //Decode the Value
                        ctrlValue = Server.UrlDecode(ctrlValue);
                        if (ctrlValue != String.Empty)
                        {
                            recipientList.Add(ctrlValue.Trim());
                        }
                        break;
                    }
                }
            }
        }

        //string pattern = "*" + User.Identity.Name + "*.pdf";
        //var dirInfo = new DirectoryInfo(@"\\10.0.0.6\Shopdata\Development\tiger\Sales Call Report\");
        //var dirInfo = new DirectoryInfo(@"C:\Sales Call Report\");
        //var file = (from f in dirInfo.GetFiles(pattern) orderby f.LastWriteTime descending select f).First();
       

        // send file
        MailMessage mailmsg = new MailMessage();
        foreach (string recipient in recipientList)
        {
            mailmsg.To.Add(recipient);
        }
        mailmsg.From = new MailAddress("office@etsdies.com");
        mailmsg.Subject = "Sales Call Report";
        
        mailmsg.Attachments.Add(new Attachment(Server.MapPath("~/Sales Call Report PDF/" + filesubPath)));
        mailmsg.Body = emailmsg;

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
                Response.Write("<SCRIPT LANGUAGE=\"JavaScript\">alert(\"Email sent!\")</SCRIPT>");
            }
            catch
            {
                hasSend = false;
                Response.Write("<SCRIPT LANGUAGE=\"JavaScript\">alert(\"Fail to send Email!\")</SCRIPT>");
            }
        }
    }

    protected void DropDownEmailList_SelectedIndexChanged(object sender, EventArgs e)
    {
        int cnt = FindOccurence("RecipientList");
        TextBox txt = new TextBox();
        txt.ID = "RecipientList-" + Convert.ToString(cnt + 1);
        txt.Text = DropDownEmailList.SelectedValue.ToString();
        pnlEmailBox.Controls.Add(txt);

        Literal lt = new Literal();
        lt.Text = "<br />";
        pnlEmailBox.Controls.Add(lt);
    }

    private int FindOccurence(string substr)
    {
        string reqstr = Request.Form.ToString();
        return ((reqstr.Length - reqstr.Replace(substr, "").Length) / substr.Length);
    }

    private void RecreateControls(string ctrlPrefix, string ctrlType)
    {
        string[] ctrls = Request.Form.ToString().Split('&');
        int cnt = FindOccurence(ctrlPrefix);
        if (cnt > 0)
        {
            Literal lt;
            for (int k = 1; k <= cnt; k++)
            {
                for (int i = 0; i < ctrls.Length; i++)
                {
                    if (ctrls[i].Contains(ctrlPrefix + "-" + k.ToString()))
                    {
                        string ctrlName = ctrls[i].Split('=')[0];
                        string ctrlValue = ctrls[i].Split('=')[1];

                        //Decode the Value
                        ctrlValue = Server.UrlDecode(ctrlValue);

                        if (ctrlType == "TextBox")
                        {
                            TextBox txt = new TextBox();
                            txt.ID = ctrlName;
                            txt.Text = ctrlValue;
                            pnlEmailBox.Controls.Add(txt);
                            lt = new Literal();
                            lt.Text = "<br />";
                            pnlEmailBox.Controls.Add(lt);
                        }

                        break;
                    }
                }
            }
        }
    }
    protected void BtnPDF_Click(object sender, EventArgs e)
    {
        // provide file to download
        HttpResponse response = HttpContext.Current.Response;
        response.AppendHeader("Content-Disposition", "attachment; filename=" + filesubPath);
        FileInfo file = new FileInfo(Server.MapPath("~/Sales Call Report PDF/" + filesubPath));
        response.AppendHeader("Content-Length", file.Length.ToString());
        response.ContentType = "application/pdf";
        response.WriteFile(Server.MapPath("~/Sales Call Report PDF/" + filesubPath));
        response.End();
    }
}