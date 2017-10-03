using System;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Security;

public partial class Account_ChangePassword : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
    }

    //protected void OnCancelButton(EventArgs e)
    //{
    //    Response.Redirect("default.aspx");
    //}

    //protected void ChangePasswordPushButton_Click(object sender, EventArgs e)
    //{
    //    if (ChangeUserPassword.CurrentPassword.Trim().Length > 0 && ChangeUserPassword.NewPassword.Trim().Length >= Membership.MinRequiredPasswordLength)
    //    {
    //        string strConnection = "server=10.0.0.6;database=tiger;user id=jamie;password=jamie;";
    //        SqlConnection sqlConnection = new SqlConnection(strConnection);
    //        // check if old password is right
    //        String query = "select [Password] from [tiger].[dbo].[user] where [UserName] like '" + HttpContext.Current.User.Identity.Name + "'";
    //        SqlCommand command = new SqlCommand(query, sqlConnection);
    //        sqlConnection.Open();
    //        SqlDataReader reader = command.ExecuteReader();
    //        if (reader.Read())
    //        {
    //            if (ChangeUserPassword.CurrentPassword.Trim() == reader[0].ToString().Trim())
    //            {
    //                query = "update [tiger].[dbo].[user] set [Password]='" + ChangeUserPassword.NewPassword + "'where [UserName] like '" + HttpContext.Current.User.Identity.Name + "'";
    //                reader.Close();
    //                command.CommandText = query;
    //                if (1 == command.ExecuteNonQuery())
    //                {
    //                    Response.Write("Password changed");
    //                }
    //                else
    //                {
    //                    Response.Write("Password changed failed");
    //                }                
    //            }
    //            else
    //            {
    //                Response.Write("Password changed failed");
    //            }
    //        }
    //        else
    //        {
    //            Response.Write("Password changed failed");
    //        }
    //        reader.Close();
    //    }
    //    else
    //    {
    //        ChangeUserPassword.ChangePasswordFailureText = "aaa";
    //        ChangeUserPassword.
    //    }

    //}
    protected void Cancel_Click(object sender, EventArgs e)
    {
        Response.Redirect("default.aspx");
    }

    protected void Update_Click(object sender, EventArgs e)
    {
        if (0 == OldPassword.Text.Length)
        {
            NewPasswordErrorLabel.Text = string.Empty;
            OldPasswordErrorLabel.Text = "Please input old password.";
        }
        else if (6 > NewPassword.Text.Length)
        {
            OldPasswordErrorLabel.Text = string.Empty;
            NewPasswordErrorLabel.Text = "New password must contain minimum 6 characters.";
        }
        else if (RepeatPassword.Text != NewPassword.Text)
        {
            OldPasswordErrorLabel.Text = string.Empty;
            NewPasswordErrorLabel.Text = "The same new password must have to be typed twice.";
        }
        else if (!ValidateOldPassword())
        {
            NewPasswordErrorLabel.Text = string.Empty;
            OldPasswordErrorLabel.Text = "Old password is incorrect.";
        }
        else
        {
            string strConnection = "server=10.0.0.6;database=tiger;user id=jamie;password=jamie;";
            SqlConnection sqlConnection = new SqlConnection(strConnection);
            String query = "update [tiger].[dbo].[user] set [Password]='" + NewPassword.Text + "'where [UserName] like '" + HttpContext.Current.User.Identity.Name + "'";
            SqlCommand command = new SqlCommand(query, sqlConnection);
            sqlConnection.Open();
            if (1 == command.ExecuteNonQuery())
            {
                Response.Redirect("ChangePasswordSuccess.aspx");
            }
        }
    }

    private bool ValidateOldPassword()
    {
        string strConnection = "server=10.0.0.6;database=tiger;user id=jamie;password=jamie;";
        SqlConnection sqlConnection = new SqlConnection(strConnection);
        // check if old password is right
        String query = "select [Password] from [tiger].[dbo].[user] where [UserName] like '" + HttpContext.Current.User.Identity.Name + "'";
        SqlCommand command = new SqlCommand(query, sqlConnection);
        sqlConnection.Open();
        SqlDataReader reader = command.ExecuteReader();
        if (reader.Read())
        {
            if (OldPassword.Text == reader[0].ToString().Trim())
            {
                return true;
            }
        }
        return false;
    }
}

/*
                 query = "update [tiger].[dbo].[user] set [Password]='" + ChangeUserPassword.NewPassword + "'where [UserName] like '" + HttpContext.Current.User.Identity.Name + "'";
            reader.Close();
            command.CommandText = query;
            if (1 == command.ExecuteNonQuery())
            {
                e.Cancel = true;
            }
            else
            {
                e.Cancel = false;
            }
        }

 */
