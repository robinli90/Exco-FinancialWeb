using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using System.Web.UI;
using System.Data.Odbc;
using System.Web.UI.WebControls;

public partial class Account_Login : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (HttpContext.Current.User.Identity.IsAuthenticated)
        {
            Response.Redirect("default.aspx");
        }
        else
        {
        }
    }

    protected void LoginUser_Authenticate(object sender, AuthenticateEventArgs e)
    {
        ((Login)sender).UserName = ((Login)sender).UserName.ToUpper();
        string strConnection = "server=10.0.0.6;database=tiger;user id=jamie;password=jamie;";
        SqlConnection sqlConnection = new SqlConnection(strConnection);
        String SQLQuery = "select count(*) from [tiger].[dbo].[user] where [UserName] like '" + ((Login)sender).UserName + "' and [Password] like '" + ((Login)sender).Password + "'";
        SqlCommand command = new SqlCommand(SQLQuery, sqlConnection);
        SqlDataReader Dr;
        sqlConnection.Open();
        Dr = command.ExecuteReader();
        if (Dr.Read())
        {
            if (Convert.ToInt16(Dr[0]) > 0)
            {
                e.Authenticated = true;
            }
            else
            {
                e.Authenticated = false;
            }
        }
        Dr.Close();
        if (e.Authenticated)
        {
            // write last logged on time
            SQLQuery = "update [tiger].[dbo].[user] set [TimeLastLoggedIn]='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' where [UserName] like '" + ((Login)sender).UserName + "' and [Password] like '" + ((Login)sender).Password + "'";
            command.CommandText = SQLQuery;
            if (1 != command.ExecuteNonQuery())
            {
                throw new Exception("write logged on time failed: " + ((Login)sender).UserName);
            }
        }
    }
}