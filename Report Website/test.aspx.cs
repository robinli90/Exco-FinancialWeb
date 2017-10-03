using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ExcoUtility;
using System.Data.Odbc;

public partial class test : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
    }
    protected void Button1_Click(object sender, EventArgs e)
    {
        string query = "select bvname from cmsdat.cust";
        ExcoODBC database = ExcoODBC.Instance;
        database.Open(Database.CMSDAT);
        OdbcDataReader reader = database.RunQuery(query);
        if (reader.Read())
        {
            TextBox1.Text = reader[0].ToString();
        }
        reader.Close();

    }
}