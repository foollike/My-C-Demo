using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text.RegularExpressions;

namespace PMWeb
{
    public partial class PMCreatAC : System.Web.UI.Page
    {
        Systemset SS = new Systemset();
        

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                DropDownList_Job.SelectedIndex = 0;
            }            
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            Label_Dis.Text = "";

            if (TextBox_Name.Text == string.Empty)
            {
                Label_Dis.Text = "姓名欄不可為空";
            }            
            else if (TextBox_Account.Text == string.Empty || !IsAlphaNumeric(TextBox_Account.Text))
            {
                Label_Dis.Text = "帳號欄不可為空或含有不法字元";
            }
            else if (TextBox_Password.Text == string.Empty || !IsAlphaNumeric(TextBox_Password.Text))
            {
                Label_Dis.Text = "密碼欄不可為空或含有不法字元";
            }
            else if (TextBox_Password.Text != TextBox_RePassword.Text || TextBox_RePassword.Text == string.Empty || !IsAlphaNumeric(TextBox_RePassword.Text))
            {
                Label_Dis.Text = "重複輸入密碼錯誤或含有不法字元";
            }            
            else if (TextBox_Account.Text.Length < 4 || TextBox_Account.Text.Length > 12)
            {
                Label_Dis.Text = "帳號長度有誤";
            }
            else if (TextBox_Password.Text.Length < 4 || TextBox_Password.Text.Length > 12)
            {
                Label_Dis.Text = "密碼長度有誤";
            }
            else if (TextBox_PSHint.Text == string.Empty)
            {
                Label_Dis.Text = "請填寫密碼提示";
            }
            else if (DropDownList_Job.SelectedIndex == 0)
            {
                Label_Dis.Text = "請選擇職位";
            }

            else
            {
                bool Alive = true; //帳號是否重覆

                try
                {
                    #region 驗證帳號是否存在

                    string[] GetMessage = SS.AccountAlive(TextBox_Account.Text, Systemset.AdPassword).Split('|');

                    if (GetMessage[0] == "Y")
                    {
                        Alive = false;
                    }
                    else
                    {
                        if (GetMessage[0] == "R")
                        {
                            Label_Dis.Text = "帳號重複";
                        }
                        else
                        {
                            Label_Dis.Text = GetMessage[1];
                        }
                    }

                    #endregion

                    #region 創建新帳號

                    if (Alive == false)
                    {
                        string Job = "其-";
                        switch (DropDownList_Job.SelectedIndex)
                        {
                            case 1:
                                Job = "企-";
                                break;
                            case 2:
                                Job = "程-";
                                break;
                            case 3:
                                Job = "美-";
                                break;
                            case 4:
                                Job = "檢-";
                                break;
                            default:
                                Job = "其-";
                                break;
                        }

                        string[] GetRegister = SS.AccountRegister(TextBox_Account.Text, TextBox_Password.Text, Job + TextBox_Name.Text, TextBox_PSHint.Text, Systemset.AdPassword).Split('|');

                        if (GetRegister[0] == "Y")
                        {
                            Response.Write("<Script language='JavaScript'>alert('註冊成功！');</Script>");
                            Server.Transfer("PMLogin.aspx");
                            Response.End();
                        }
                        else
                        {
                            Label_Dis.Text = GetRegister[1];
                        }
                    }

                    #endregion
                }
                catch
                {
                    Label_Dis.Text = "註冊失敗";
                }

            }

        }

        /// 檢查是否為數字或者字母
        public static bool IsAlphaNumeric(String InputString)
        {
            return (!Regex.IsMatch(InputString, "[^a-zA-Z0-9]")) ? true : false;
        }

        public static bool IsEmailAddress(String InputString)
        {
            return (Regex.IsMatch(InputString, @"^\w+((-\w+)|(\.\w+))*\@\w+((\.|-)\w+)*\.\w+$")) ? true : false;
        }

        protected void Button_Back_Click(object sender, EventArgs e)
        {
            Response.Redirect("PMLogin.aspx");
        }
    }
}