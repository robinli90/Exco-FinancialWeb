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
    private string fiscalYear, plant, desc, sort_option;
    private string fromYear, toYear;
    private string[,] po_info = new string[45000, 6];
    private string plant_NO;

    // Return the query for the "ALL" selection
    public string Get_Query(bool has_desc, string plant, string startYear, string endYear, string sorttype="date", string desc="%")
    {
        string start_date = "'" + startYear + "-10-01'";
        string end_date = "'" + endYear + "-09-30'";
        string with_desc = "and upper(c.jrdes1) like upper('%" + desc + "%') " ;
        string asc_desc = "asc";
        if (sorttype == "unit_price") // Automatically sort by highest price because it makes sense to do so this way
        {
            asc_desc = "desc";
        }
        if (!has_desc)
        {
            with_desc = "";
        }

        return  "select partno, unit_price, unit, date, C.JRDES1 as DESCRIPTION, C.JRVND# from " +
                "(select a.KBVPT# as PartNO, a.KBUPRC as UNIT_Price, a.KBPUNT as UNIT, a.KBRDAT as " +
                "DATE, a.KBPLNT as PLANT from cmsdat.poi as a) as A, cmsdat.poptvn as c where partno=c.jrvpt# and partno <> '' " +
                with_desc +
                "and plant = '" + plant + "' " +
                "and date between " + start_date + " and " + end_date + " order by " + sorttype + " " + asc_desc;
    }

    protected void Page_Load(object sender, EventArgs e)
    {
    }

    protected void ButtonGenerate_Click(object sender, EventArgs e)
    {

        ReportGrid.Visible = true;
        DataTable report = new DataTable();


        // Description Toggle Query
        desc = desc_box1.Text.ToString();
        bool has_descrip = true;
        if (desc == "")
        {
            has_descrip = false;
        }

        // Sorting Query
        sort_option = sortby.SelectedValue.ToString();
        string sort_option2 = "";

        if (sort_option == "Date")
        {
            sort_option2 = "date";
        }
        else if (sort_option == "Vendor")
        {
            sort_option2 = "jrvnd#";
        }
        else if (sort_option == "Part")
        {
            sort_option2 = "partno";
        }
        else if (sort_option == "Description")
        {
            sort_option2 = "description";
        }
        else if (sort_option == "Price")
        {
            sort_option2 = "unit_price";
        }
        

        // Year Query 
        fiscalYear = YearChoice1.SelectedValue.ToString();
        bool all_years = false;
        if (fiscalYear == "All" || fiscalYear == "ALL")
        {
            fromYear = "2010";
            toYear = "2017";
            all_years = true;
        }
        else
        {
            fromYear = (Convert.ToInt32(fiscalYear) - 1).ToString();
            toYear = fiscalYear;
        }

        // Plant Query
        plant = PlantList1.SelectedValue.ToString();
        if (plant == "Markham")
        {
            plant_NO = "001";
        }
        else if (plant == "Michigan")
        {
            plant_NO = "003";
        }
        else if (plant == "Texas")
        {
            plant_NO = "005";
        }
        else if (plant == "Colombia")
        {
            plant_NO = "004";
        }

        var connection = new OdbcConnection(ConfigurationManager.ConnectionStrings["cms1"].ConnectionString);
        connection.Open();
        var command = new OdbcCommand();
        command.Connection = connection;
        var query = Get_Query(has_descrip, plant_NO, fromYear, toYear, sort_option2, desc);

        GridView gridView = ReportGrid;
        if (all_years)
        {
            gridView.Caption = "<font size=\"5\" color=\"red\"> Vendor Parts Listing in all years</font><br/></br>";
        }
        else
        {
            gridView.Caption = "<font size=\"5\" color=\"red\"> Vendor Parts Listing in " + toYear.ToString() + "</font><br/></br>";
        }
        //gridView.Caption = "<font size=\"5\" color=\"red\"> Query: " + query + "</font>";

        command.CommandText = query;
        var reader = command.ExecuteReader();
        int iteration = 0;
        while (reader.Read())// && iteration < 5000)
        {
            po_info[iteration, 0] = reader[0].ToString().Trim(); // Part NO
            po_info[iteration, 1] = reader[1].ToString().Trim(); // Unit Price
            po_info[iteration, 2] = reader[2].ToString().Trim(); // Unit
            po_info[iteration, 3] = reader[3].ToString().Trim(); // Date
            po_info[iteration, 4] = reader[4].ToString().Trim(); // Description
            po_info[iteration, 5] = reader[5].ToString().Trim(); // Vendor NO
            iteration++;
        }
        reader.Close();
        connection.Close();


        report.Columns.Add("Vendor");
        report.Columns.Add("Part #");
        report.Columns.Add("Description");
        report.Columns.Add("Date");
        report.Columns.Add("Unit Price");
        report.Columns.Add("Unit");
        List<DataRow> rowList = new List<DataRow>();
        for (int i = 0; i < iteration; i++)
        {
            DataRow row = report.NewRow();    
            row["Vendor"] = po_info[i, 5];
            row["Part #"] = po_info[i, 0];
            row["Description"] = po_info[i, 4];
            row["Date"] = po_info[i, 3].Substring(0, po_info[i, 3].Length - 11);
            row["Unit Price"] = po_info[i, 1];
            row["Unit"] = po_info[i, 2];
            report.Rows.Add(row);
        }

        gridView.DataSource = report;
        gridView.DataBind();
        po_info = new string[45000, 6];

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