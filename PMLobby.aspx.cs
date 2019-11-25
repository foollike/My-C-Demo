using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;

namespace PMWeb
{
    public partial class PMLobby1 : System.Web.UI.Page
    {
        Systemset SS = new Systemset();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                #region 專案選單初始

                DataSet PMenu = new DataSet();
                PMenu = SS.GetSqlTable("SELECT ProjectID, ProjectName, Description, CreatDate, UpdateDate FROM PMProject WHERE (Alive = '1') ORDER BY ProjectID");

                try
                {
                    if (PMenu.Tables[0].Rows.Count > 0)
                    {
                        Button_DateCheck.Enabled = true;
                        Button_MList.Enabled = true;
                        TextBox_DatePick.Enabled = true;

                        for (int i = 0; i < PMenu.Tables[0].Rows.Count; i++)
                        {
                            ListItem items = new ListItem();
                            items.Text = PMenu.Tables[0].Rows[i]["ProjectName"].ToString();
                            items.Value = PMenu.Tables[0].Rows[i]["ProjectID"].ToString();
                            DropDownList_Project.Items.Add(items);
                        }
                    }
                    else
                    {
                        Button_DateCheck.Enabled = false;
                        Button_MList.Enabled = false;
                        TextBox_DatePick.Enabled = false;

                        ListItem items = new ListItem();

                        if (Session["SPermission"].ToString() == "1")
                        {
                            items.Text = "請至管理者頁面創建專案";
                        }
                        else
                        {
                            items.Text = "請聯絡管理者增加專案";
                        }
                        DropDownList_Project.Items.Add(items);
                        DropDownList_Project.Enabled = false;
                    }
                }

                catch
                {
                    Relogin(1, 0);
                }

                #endregion

                #region 初始化

                TextBox_DatePick.Text = DateTime.Now.ToString("yyyy-MM");
                Session["PPermission"] = "0"; //初始化專案管理

                #endregion

                if (Session["UserID"] != null) //確認登入成功
                {
                    //Response.Write("<Script language='JavaScript'>alert('" + Session["Name"].ToString() + " 歡迎回來!');</Script>");
                    NameDis.Text = "使用者 : " + Session["Name"].ToString().Substring(2, Session["Name"].ToString().Length - 2);

                    //確認使用者是否為最高權限管理員，如是可以跳過一切檢察

                    //SPermission : 系統管理員權限 0=一般使用者 1=系統管理員
                    //PPermission : 專案管理員權限 0=一般使用者 1=專案管理員

                    if (Session["SPermission"].ToString() == "1")
                    {
                        ShowADButton(1);

                        if (PMenu.Tables[0].Rows.Count > 0)
                        {
                            ShowCalender();
                        }
                        else
                        {
                            GridView_All.Visible = false;
                        }
                    }
                    else
                    {
                        try
                        {
                            string Tsql = "SELECT ProjectID, Permission FROM PMPMember WHERE (UserID = '" + Session["UserID"].ToString() + "') AND (Alive = '1') ORDER BY Permission DESC, ProjectID";

                            //檢查PMPMember表(專案成員)，使用者是否有加入任何專案。
                            string[] AdCheck = SS.GetSqlString(Tsql, "ProjectID,Permission").Split(',');

                            //如未加入任何專案，所有使用者選項都會被鎖起來
                            if (AdCheck[0] != string.Empty)
                            {
                                //DropDownList_Project.SelectedItem.Value = DropDownList_Project.Items.FindByValue(AdCheck[0]).ToString();//切換到擁有的專案 
                                DropDownList_Project.SelectedIndex = DropDownList_Project.Items.IndexOf(DropDownList_Project.Items.FindByValue(AdCheck[0]));

                                if (AdCheck[1] == "1") //檢查是否為專案管理員
                                {
                                    Session["PPermission"] = "1";
                                }

                                ShowADButton(1);

                                if (PMenu.Tables[0].Rows.Count > 0)
                                {
                                    ShowCalender();
                                }
                                else
                                {
                                    GridView_All.Visible = false;
                                }
                            }
                            else
                            {
                                Relogin(0, 1);
                                ShowADButton(2);                      
                            }
                        }
                        catch
                        {
                            Relogin(1, 0);
                        }
                    }
                }
                else
                {
                    Relogin(1, 2);
                }
            }
        }

        private void ParmissionCheck(string ProjectID) //檢查使用者專案權限
        {
            //檢查PMPMember表(專案成員)，使用者是否為專案管理員。
            string[] AdCheck = SS.GetSqlString("SELECT Permission FROM PMPMember WHERE (UserID = '" + Session["UserID"].ToString() + "') AND (ProjectID = '" + ProjectID + "') AND (Alive = '1')", "Permission").Split(',');

            //是系統管理員可跳過檢察
            if (Session["SPermission"].ToString() == "1")
            {
                ShowADButton(1);
            }
            else
            {
                if (AdCheck[0] != string.Empty)
                {
                    Session["PPermission"] = AdCheck[0];
                    ShowADButton(1);
                }
                else
                {
                    ShowADButton(2);
                    Relogin(0, 3);
                }
            }
        }

        private void ShowADButton(int kind) //檢查是否開啟管理者介面
        {
            try
            {
                switch (kind)
                {
                    case 1:
                        if (Session["PPermission"].ToString() == "1" || Session["SPermission"].ToString() == "1") //只要是系統或是專案其中之一為管理員即可開啟
                        {
                            Button_MCreate.Visible = true;
                            Button_Administrator.Visible = true;
                            Button_MList.Visible = true;
                            GridView_All.Visible = true;
                            TextBox_DatePick.Visible = true;
                            Label_Date.Visible = true;
                            Button_DateCheck.Visible = true;
                            Panel_Hint.Visible = true;
                        }
                        else //為普通組員
                        {
                            Button_MCreate.Visible = false;
                            Button_Administrator.Visible = false;
                            Button_MList.Visible = true;
                            GridView_All.Visible = true;
                            TextBox_DatePick.Visible = true;
                            Label_Date.Visible = true;
                            Button_DateCheck.Visible = true;
                            Panel_Hint.Visible = true;
                        }
                        break;

                    case 2: //沒有開啟專案的權限

                        Button_MCreate.Visible = false;
                        Button_Administrator.Visible = false;
                        Button_MList.Visible = false;
                        GridView_All.Visible = false;
                        TextBox_DatePick.Visible = false;
                        Label_Date.Visible = false;
                        Button_DateCheck.Visible = false;
                        Panel_Hint.Visible = false;

                        break;
                }

                #region 是否開啟創建任務按鈕

                string Tsql = "SELECT UserID, MemberName FROM PMPMember WHERE (ProjectID = '" + DropDownList_Project.SelectedValue.ToString() + "') AND (Permission = '0') AND (Alive = '1')";

                string HaveProjectMember = SS.GetSqlString(Tsql, "UserID,MemberName");

                if (HaveProjectMember == string.Empty)
                {
                    Button_MCreate.Enabled = false;
                }
                else
                {
                    Button_MCreate.Enabled = true;
                }

                #endregion

            }
            catch
            {
                Relogin(1, 0);
            }
        }               

        private int GetDaysInMonth(int month, int year) // 讀取一個月有多少天
        {
            int days = 0;
            switch (month)
            {
                case 1:
                case 3:
                case 5:
                case 7:
                case 8:
                case 10:
                case 12:
                    days = 31;
                    break;
                case 4:
                case 6:
                case 9:
                case 11:
                    days = 30;
                    break;
                case 2:
                    if (((year % 4) == 0) && ((year % 100) != 0) || ((year % 400) == 0))
                        days = 29;
                    else
                        days = 28;
                    break;
            }
            return days;
        }

        protected void DropDownList_Project_SelectedIndexChanged(object sender, EventArgs e) //專案切換
        {
            ParmissionCheck(DropDownList_Project.SelectedValue.ToString());
            ShowCalender();         
        }

        private void ShowCalender() //顯示行事曆
        {            
            try
            {
                Session["PickDate"] = TextBox_DatePick.Text; //取得目前日期                
                string[] PickDate = Session["PickDate"].ToString().Split('-');

                #region 顯示日期

                Session["Days"] = 0; 

                int days = GetDaysInMonth(Convert.ToInt32(PickDate[1]), Convert.ToInt32(PickDate[0])); //依照天數創建表格
                
                DataTable dt = new DataTable();
                DataRow drow;
                //DataColumn dcol1 = new DataColumn("Column1", typeof(string));
                //dt.Columns.Add(dcol1);

                for (int i = 1; i <= days; i++)
                {
                    drow = dt.NewRow();
                    dt.Rows.Add(drow);
                }

                #endregion

                DataSet CalenderDs = new DataSet();

                string Tsql = "";

                if (Session["PPermission"].ToString() == "1" || Session["SPermission"].ToString() == "1")
                {                    
                    for (int i = 1; i <= days; i++)
                    {
                        Tsql += "SELECT ProjectID, MissionID, MType, Creator, Title, Description, StartDate AS StartDate, EndDate, CreatDate, ResponsibleName, Complete, CompleteDate FROM PMMission WHERE (ProjectID = '" + DropDownList_Project.SelectedValue.ToString() + "') AND (DATEPART(MM, StartDate) = '" + PickDate[1] + "') AND (DATEPART(yyyy, StartDate) = '" + PickDate[0] + "') AND (DATEPART(dd, StartDate) = '" + i + "') AND (Alive = '1');";
                    }  
                }
                else
                {
                    #region 取得使用者擁有的任務代號

                    string LTsql = "SELECT MissionID FROM PMMMember WHERE (ProjectID = '" + DropDownList_Project.SelectedValue.ToString() + "') AND (UserID = '" + Session["UserID"].ToString() + "') AND (Alive = '1') ORDER BY MissionID";
                    DataSet MMerberList = SS.GetSqlTable(LTsql);
                    string UserMissionID = "";

                    if (MMerberList.Tables[0].Rows.Count > 0)
                    {
                        UserMissionID = "AND (MissionID = '" + MMerberList.Tables[0].Rows[0]["MissionID"].ToString() + "'";

                        for (int i = 1; i < MMerberList.Tables[0].Rows.Count; i++)
                        {
                            UserMissionID += " OR MissionID = '" + MMerberList.Tables[0].Rows[i]["MissionID"].ToString() + "'";
                        }

                        UserMissionID += ")";

                        for (int i = 1; i <= days; i++)
                        {
                            Tsql += "SELECT ProjectID, MissionID, MType, Creator, Title, Description, StartDate AS StartDate, EndDate, CreatDate, ResponsibleName, Complete, CompleteDate FROM PMMission WHERE (ProjectID = '" + DropDownList_Project.SelectedValue.ToString() + "') AND (DATEPART(MM, StartDate) = '" + PickDate[1] + "') AND (DATEPART(yyyy, StartDate) = '" + PickDate[0] + "') AND (DATEPART(dd, StartDate) = '" + i + "') AND (Alive = '1')" + UserMissionID + ";";
                        }
                    }                    

                    #endregion                    
                }

                CalenderDs = SS.GetSqlTable(Tsql);
                Session["CalenderDs"] = CalenderDs;    

                GridView_All.DataSource = dt;
                GridView_All.DataBind(); 
            }
            catch
            {
                Relogin(1, 0);
            }


        }

        protected void Button_Logout_Click(object sender, EventArgs e) //登出按鈕
        {            
            Response.Redirect("PMLogin.aspx");
        }                
                
        protected void GridView_All_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                try
                {
                    #region 顯示日期
                    Session["Days"] = Convert.ToInt32(Session["Days"]) + 1; //計算天數
                    string[] PickDate = Session["PickDate"].ToString().Split('-');

                    #region 取得星期幾

                    DateTime dt = new DateTime(Convert.ToInt32(PickDate[0]), Convert.ToInt32(PickDate[1]), Convert.ToInt32(Session["Days"]));
                    string DayOfWeek = "";
                    switch (dt.DayOfWeek.ToString("d"))
                    {
                        case "1":
                            DayOfWeek = "一";
                            break;
                        case "2":
                            DayOfWeek = "二";
                            break;
                        case "3":
                            DayOfWeek = "三";
                            break;
                        case "4":
                            DayOfWeek = "四";
                            break;
                        case "5":
                            DayOfWeek = "五";
                            break;
                        case "6":
                            DayOfWeek = "六";
                            break;
                        case "0":
                            DayOfWeek = "日";
                            break;
                    }
                    #endregion

                    Label Label_DateShow;
                    Label_DateShow = (Label)e.Row.Cells[0].FindControl("Label_DateShow");
                    Label_DateShow.Text = PickDate[1] + " / " + Session["Days"].ToString() + " " + "<br />" + "星期" + DayOfWeek;
                    #endregion
                }
                catch
                {
                    Relogin(0, 4);
                }

                try
                {
                    #region 創建子行事曆

                    DataTable dtable = new DataTable();
                    DataRow drow;
                    Session["CalenderIndex"] = 0; //定位顯示該天的哪一筆工作

                    if (Session["CalenderDs"] != null)
                    {
                        DataSet ds = (DataSet)Session["CalenderDs"];

                        for (int i = 0; i < ds.Tables[Convert.ToInt32(Session["Days"]) - 1].Rows.Count; i++) //確認要創建幾行工作列
                        {
                            drow = dtable.NewRow();
                            dtable.Rows.Add(drow);
                        }

                        GridView GridView_Dis = (GridView)e.Row.Cells[0].FindControl("GridView_Dis");

                        GridView_Dis.DataSource = dtable;
                        GridView_Dis.DataBind();
                    }                    

                    #endregion
                }
                catch
                {
                    Relogin(0, 5);
                }
            }           
        }              

        protected void Button_DateCheck_Click(object sender, EventArgs e) //確認日期按鈕
        {
            ShowCalender();
        }

        protected void GridView_Dis_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                try
                {
                    if (Session["CalenderDs"] != null)
                    {
                        DataSet MainDs = (DataSet)Session["CalenderDs"]; //使用在ShowCalender抓好的表來顯示資料(一天一個表整合在一個Dataset裡)

                        //DateTime StartDate = Convert.ToDateTime(MainDs.Tables[Convert.ToInt32(Session["Days"]) - 1].Rows[Convert.ToInt32(Session["CalenderIndex"])]["StartDate"].ToString());
                        DateTime EndDate = Convert.ToDateTime(MainDs.Tables[Convert.ToInt32(Session["Days"]) - 1].Rows[Convert.ToInt32(Session["CalenderIndex"])]["EndDate"].ToString());

                        Label Label_DType = (Label)e.Row.Cells[0].FindControl("Label_DType");
                        Label Label_DTitle = (Label)e.Row.Cells[1].FindControl("Label_DTitle");
                        Label Label_DEndDate = (Label)e.Row.Cells[2].FindControl("Label_DEndDate");
                        Label Label_DDuration = (Label)e.Row.Cells[3].FindControl("Label_DDuration");
                        Label Label_DMember = (Label)e.Row.Cells[4].FindControl("Label_DMember");
                        HyperLink HyperLink_DetailLink = (HyperLink)e.Row.Cells[5].FindControl("HyperLink_DetailLink");

                        string[] TypeName = SS.GetMissionType();
                        Label_DType.Text = TypeName[(Convert.ToInt32(MainDs.Tables[Convert.ToInt32(Session["Days"]) - 1].Rows[Convert.ToInt32(Session["CalenderIndex"])]["MType"].ToString()))]; //顯示類型                         
                        Label_DTitle.Text = MainDs.Tables[Convert.ToInt32(Session["Days"]) - 1].Rows[Convert.ToInt32(Session["CalenderIndex"])]["Title"].ToString(); //顯示標題
                        Label_DEndDate.Text = EndDate.ToString("yyyy-MM-dd"); //顯示工作結束時間

                        string Deadline = "";

                        if (MainDs.Tables[Convert.ToInt32(Session["Days"]) - 1].Rows[Convert.ToInt32(Session["CalenderIndex"])]["Complete"].ToString() == "1")
                        {
                            Deadline = "已完成";
                            Label_DEndDate.ForeColor = System.Drawing.Color.Green;
                            Label_DDuration.ForeColor = System.Drawing.Color.Green;
                        }
                        else
                        {
                            if ((EndDate - DateTime.Now).Days < 1)
                            {
                                Deadline = "已截止";
                                Label_DEndDate.ForeColor = System.Drawing.Color.Red;
                                Label_DDuration.ForeColor = System.Drawing.Color.Red;
                            }
                            else
                            {
                                Deadline = (EndDate - DateTime.Now).Days.ToString() + "天";
                            }
                        }

                        Label_DDuration.Text = Deadline; //顯示工作期限天數                    
                        Label_DMember.Text = MainDs.Tables[Convert.ToInt32(Session["Days"]) - 1].Rows[Convert.ToInt32(Session["CalenderIndex"])]["ResponsibleName"].ToString(); //顯示負責人

                        int isAdministrator = 0;
                        if (Session["PPermission"].ToString() == "1" || Session["SPermission"].ToString() == "1")
                        {
                            isAdministrator = 1;
                        }

                        HyperLink_DetailLink.NavigateUrl = "javascript:void window.open('MissionDetail.aspx?ProjectID=" + MainDs.Tables[Convert.ToInt32(Session["Days"]) - 1].Rows[Convert.ToInt32(Session["CalenderIndex"])]["ProjectID"].ToString() + "&MissionID=" + MainDs.Tables[Convert.ToInt32(Session["Days"]) - 1].Rows[Convert.ToInt32(Session["CalenderIndex"])]["MissionID"].ToString() + "&isAdministrator=" + isAdministrator.ToString() + "&UserID=" + Session["UserID"].ToString() + "&Name=" + Session["Name"].ToString() + "','NewWindow',config='height=780,width=530,left=800,top=200','menubar=no','status=no','scrollbars=yes');"; //詳細資料連結

                        Session["CalenderIndex"] = Convert.ToInt32(Session["CalenderIndex"]) + 1; //累計行數
                    }                                     
                }
                catch
                {
                    Relogin(0, 5);
                }
            }
        }

        protected void Relogin(int NeedRelogin, int ErrorCode) //Error應對
        {
            string Message = "";

            #region 錯誤訊息

            switch (ErrorCode)
            {
                case 0:
                    Message = "系統錯誤,請重新登入!";
                    break;
                case 1:
                    Message = "需加入最少一個專案才能使用本系統喔";
                    break;
                case 2:
                    Message = "登入失敗,請重新登入!";
                    break;
                case 3:
                    Message = "沒有觀看此專案的權限";
                    break;
                case 4:
                    Message = "顯示日期錯誤";
                    break;
                case 5:
                    Message = "行事曆創建失敗";
                    break;
            }
            #endregion

            Response.Write("<script> alert('" + Message + "');</script>");

            if (NeedRelogin == 1)
            {
                Server.Transfer("PMLogin.aspx");
                Response.End();
            }
        }

        protected void Button_MCreate_Click(object sender, EventArgs e)
        {
            Response.Redirect("MissionCreat.aspx?ProjectID=" + DropDownList_Project.SelectedValue.ToString() + "&UserID=" + Session["UserID"].ToString() + "&Name=" + Session["Name"].ToString());
        }

        protected void Button_Administrator_Click(object sender, EventArgs e)
        {
            Response.Redirect("PMAdministrator.aspx?UserID=" + Session["UserID"].ToString() + "&Name=" + Session["Name"].ToString() + "&SPermission=" + Session["SPermission"].ToString());
        }

        protected void Button_MList_Click(object sender, EventArgs e)
        {
            int isAdministrator = 0;
            if (Session["PPermission"].ToString() == "1" || Session["SPermission"].ToString() == "1")
            {
                isAdministrator = 1;
            }
            Response.Redirect("MissionList.aspx?ProjectID=" + DropDownList_Project.SelectedValue + "&ProjectName=" + DropDownList_Project.SelectedItem + "&UserID=" + Session["UserID"].ToString() + "&isAdministrator=" + isAdministrator.ToString());
        }
    }
}