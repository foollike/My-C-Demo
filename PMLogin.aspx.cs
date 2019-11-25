using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text.RegularExpressions;

namespace PMWeb
{
    public partial class PMLogin : System.Web.UI.Page
    {
        Systemset SS = new Systemset();
        string AdPassword = "h27418221";

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button_Login_Click(object sender, EventArgs e)
        {
            #region 確認玩家帳號

            try
            {
                if (TextBox_Account.Text == string.Empty || !IsAlphaNumeric(TextBox_Account.Text))
                {
                    Label_Dis.Text = "帳號欄不可為空，或不法字元";
                }
                else if (TextBox_Password.Text == string.Empty || !IsAlphaNumeric(TextBox_Password.Text))
                {
                    Label_Dis.Text = "密碼欄不可為空，或不法字元";
                }
                else
                {
                    string[] AccountAcess = SS.AcountAccess(TextBox_Account.Text, TextBox_Password.Text, AdPassword).Split('|');

                    if (AccountAcess[0] == "Y")
                    {
                        Session["UserID"] = AccountAcess[1];
                        Session["UserAccount"] = AccountAcess[2];
                        Session["UserPassword"] = AccountAcess[3];
                        Session["Name"] = AccountAcess[4];
                        Session["SPermission"] = AccountAcess[5];

                        //Response.Write("<Script language='JavaScript'>alert('" + AccountAcess[4] + " 歡迎回來!');</Script>");
                        Response.Redirect("PMLobby.aspx");
                        //Server.Transfer("PMLobby.aspx");              
                    }
                    else
                    {
                        Label_Dis.Text = AccountAcess[1];
                    }
                }
            }
            catch
            {
                Label_Dis.Text = "系統錯誤，請稍候再試";
            }

            #endregion
        }

        /// 檢查是否為數字或者字母
        public static bool IsAlphaNumeric(String InputString)
        {
            return (!Regex.IsMatch(InputString, "[^a-zA-Z0-9]")) ? true : false;
        }

        protected void Button_CreatAC_Click1(object sender, EventArgs e)
        {
            Server.Transfer("PMCreatAC.aspx");
        }

        protected void LinkButton_PWHint_Click(object sender, EventArgs e)
        {
            try
            {
                if (TextBox_Account.Text != string.Empty)
                {
                    string[] GetMessage = SS.AccountAlive(TextBox_Account.Text, AdPassword).Split('|');

                    if (GetMessage[0] == "R")
                    {
                        if (GetMessage[2] == "1")
                        {
                            Label_Dis.Text = "此為管理者帳號，不顯示提示";
                        }
                        else
                        {
                            Label_Dis.Text = GetMessage[3];
                        }                        
                    }
                    else
                    {
                        if (GetMessage[0] == "Y")
                        {
                            Label_Dis.Text = "無此帳號";
                        }
                        else
                        {
                            Label_Dis.Text = GetMessage[1];
                        }       
                    }
                }
                else
                {
                    Label_Dis.Text = "請先輸入帳號";
                }
            }
            catch
            {
                Label_Dis.Text = "系統錯誤，請稍候再試";
            }
        }
    }
}